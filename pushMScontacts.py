import requests
import time
from tqdm import tqdm
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed

# IT Glue API rate limits
MAX_REQUESTS_PER_SEC = 9
DELAY = 1.0 / MAX_REQUESTS_PER_SEC
MAX_RETRIES = 5


# ------------------------
#  Robust request handler
# ------------------------
def safe_request(method, url, headers=None, json=None, params=None, max_retries=5):
    for attempt in range(max_retries):
        try:
            resp = requests.request(method, url, headers=headers, json=json, params=params)

            # If success, return immediately
            if resp.status_code < 400:
                return resp

            # If Autotask returns 500, PRINT THE REAL MESSAGE
            if resp.status_code == 500:
                try:
                    print(f"[AT ERROR BODY] {resp.json()}")
                except:
                    print(f"[AT ERROR RAW] {resp.text}")

                # retry loop continues
        except Exception as e:
            print(f"[EXCEPTION] {e}")

        # Retry with backoff
        wait = 2 ** attempt
        print(f"[RETRY] {method} {url} failed ({resp.status_code if 'resp' in locals() else 'NO RESP'}). Retrying in {wait}s...")
        time.sleep(wait)

    # After max retries
    print(f"[ERROR] Max retries reached for {url}")
    return None


# ------------------------
# Original logic continues
# ------------------------
def contact_syncs_with(contact_data, target_adapter):
    for item in contact_data.get("included", []):
        attributes = item.get("attributes", {})
        if attributes.get("adapter-type-name") == target_adapter and attributes.get("sync"):
            return True
    return False


def get_all_contact_ids(api_key, org_id):
    headers = {"x-api-key": api_key, "Content-Type": "application/vnd.api+json"}
    contact_ids = []
    url = f"https://api.eu.itglue.com/organizations/{org_id}/relationships/contacts?page[size]=100"

    while url:
        resp = safe_request("GET", url, headers=headers)
        if not resp or resp.status_code != 200:
            tqdm.write(f"[WARNING] Failed to fetch contact IDs for org {org_id}: {resp.status_code if resp else 'N/A'}")
            break
        data = resp.json()
        contact_ids.extend(item["id"] for item in data.get("data", []))
        url = data.get("links", {}).get("next")

    return contact_ids


def fetch_contact_details(contact_id, org_id, headers):
    time.sleep(DELAY)
    url = f"https://api.eu.itglue.com/organizations/{org_id}/relationships/contacts/{contact_id}?include=adapters_resources,contact_methods,related_items"
    resp = safe_request("GET", url, headers=headers)
    return contact_id, resp


def contact_has_ms_license(contact_data):
    for item in contact_data.get("included", []):
        if item.get("type") == "tags":
            attrs = item.get("attributes", {})
            if attrs.get("resource-type-name") == "Microsoft Licenses":
                return True
    return False


def get_microsoft_only_contacts(api_key, org_id, licensed=True):
    headers = {"x-api-key": api_key, "Content-Type": "application/vnd.api+json"}
    contact_ids = get_all_contact_ids(api_key, org_id)
    ms_only_contacts = []

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = [executor.submit(fetch_contact_details, cid, org_id, headers) for cid in contact_ids]
        for future in tqdm(as_completed(futures), total=len(contact_ids), desc=f"Filtering contacts for org {org_id}"):
            contact_id, contact_resp = future.result()
            if not contact_resp or contact_resp.status_code != 200:
                print(f"[WARNING] Failed to fetch contact {contact_id}")
                continue

            contact_data = contact_resp.json()
            syncs_with_microsoft = contact_syncs_with(contact_data, "Microsoft")
            syncs_with_autotask = contact_syncs_with(contact_data, "Autotask")
            has_license = contact_has_ms_license(contact_data)

            # Apply licensed/unlicensed filter
            if syncs_with_microsoft and not syncs_with_autotask and ((licensed and has_license) or (not licensed and not has_license)):
                ms_only_contacts.append((contact_id, contact_data))

    return ms_only_contacts


def extract_emails_and_phones(contact_data):
    emails = set()
    phones = set()
    attrs = contact_data.get("data", {}).get("attributes", {})

    for email_entry in attrs.get("contact-emails", []):
        val = email_entry.get("value")
        if val:
            emails.add(val.strip())

    for phone_entry in attrs.get("contact-phones", []):
        val = phone_entry.get("value")
        if val:
            phones.add(val.strip())

    for field in ["emailAddress", "emailAddress2", "emailAddress3"]:
        val = attrs.get(field)
        if val:
            emails.add(val.strip())

    for field in ["phone", "mobilePhone", "alternatePhone", "faxNumber", "extension"]:
        val = attrs.get(field)
        if val and val.strip().lower() != "n/a":
            phones.add(val.strip())

    for item in contact_data.get("included", []):
        if item.get("type") == "contact_methods":
            attr = item.get("attributes", {})
            label = attr.get("label", "").lower()
            value = attr.get("value", "")
            if value:
                if "email" in label:
                    emails.add(value.strip())
                elif any(x in label for x in ["phone", "mobile", "fax"]):
                    phones.add(value.strip())

    return list(emails), list(phones)


def get_autotask_remote_id_from_included(included):
    for item in included:
        attrs = item.get("attributes", {})
        if attrs.get("adapter-type-name") == "Autotask":
            return attrs.get("remote-id")
    return None


# Fetch all Autotask contacts once per company (with retry)
def get_existing_autotask_contacts(base_url, headers, company_id):
    url = f"{base_url}/Companies/{company_id}/Contacts"
    existing_emails = set()

    while url:
        resp = safe_request("GET", url, headers=headers)
        if not resp or resp.status_code != 200:
            print(f"[WARNING] Failed to fetch existing contacts for company {company_id}")
            break

        data = resp.json()
        for item in data.get("items", []):
            email = (item.get("emailAddress") or "").strip().lower()
            if email:
                existing_emails.add(email)

        next_link = data.get("pageDetails", {}).get("nextPageUrl")
        url = next_link if next_link else None

    return existing_emails


def create_contact_in_autotask(base_url, auth_headers, company_id, contact_data, existing_emails_cache):
    email = (contact_data.get("email") or "").strip().lower()

    # Skip if email already exists
    if email and email in existing_emails_cache:
        print(f"[SKIP] Contact already exists in Autotask: {email}")
        return False

    first = contact_data.get("firstName") or ""
    last = contact_data.get("lastName") or ""
    phone = contact_data.get("phone") or ""

    # Autotask requires last name
    if not last:
        print(f"[SKIP] Missing last name for contact '{first} {last}', email '{email}'. Autotask requires a last name.")
        return False

    # Autotask requires email
    if not email:
        print(f"[SKIP] Missing email for contact '{first} {last}'.")
        return False

    #  Autotask Contact JSON
    payload = {
        "IsActive": 1,
        "FirstName": first,
        "LastName": last,
        "EmailAddress": email,
        "Phone": phone
    }

    url = f"{base_url}/Companies/{company_id}/Contacts"

    resp = safe_request("POST", url, headers=auth_headers, json=payload)

    if not resp:
        print(f"[ERROR] Failed to create contact {first} {last}: No response")
        return False

    if resp.status_code != 200:
        try:
            err = resp.json()
        except:
            err = resp.text
        print(f"[ERROR] Failed to create contact {first} {last}: Status {resp.status_code}, Response: {err}")
        return False

    existing_emails_cache.add(email)
    return True


def get_autotask_syncing_orgs(api_key):
    headers = {"x-api-key": api_key, "Content-Type": "application/vnd.api+json"}
    base_url = "https://api.eu.itglue.com/organizations"
    page = 1
    all_orgs = []
    autotask_orgs = []

    print("Fetching organizations...")

    while True:
        url = f"{base_url}?page[number]={page}"
        resp = safe_request("GET", url, headers=headers)
        if not resp or resp.status_code != 200:
            print(f"[ERROR] Failed to fetch orgs (page {page})")
            break

        orgs = resp.json().get("data", [])
        if not orgs:
            break
        all_orgs.extend(orgs)
        page += 1

    print("Checking for Autotask sync status...")

    for org in tqdm(all_orgs, desc="Checking organizations", unit="org"):
        org_id = org["id"]
        org_name = org["attributes"]["name"]
        detail_url = f"{base_url}/{org_id}/?include=adapters_resources"
        detail_resp = safe_request("GET", detail_url, headers=headers)
        if not detail_resp or detail_resp.status_code != 200:
            print(f"[WARN] Failed to check org {org_name}")
            continue

        included = detail_resp.json().get("included", [])
        for item in included:
            attributes = item.get("attributes", {})
            if (
                attributes.get("adapter-type-name") == "Autotask"
                and attributes.get("sync")
                and not attributes.get("orphaned")
            ):
                autotask_orgs.append((org_id, org_name, included))
                break

    return autotask_orgs


def main():
    api_key = input("Enter your IT Glue API key: ").strip()
    user_name = input("Enter your AT username: ").strip()
    secret = input("Enter your AT secret: ").strip()
    integration_code = input("Enter your AT APIIntegrationCode: ").strip()

    # Prompt to choose licensed or unlicensed Microsoft contacts
    licensed_input = input("Do you want licensed or unlicensed Microsoft contacts? [licensed/unlicensed]: ").strip().lower()
    licensed = licensed_input == "licensed"

    # Optional exclusion prompts
    exclude_ids_input = input("Enter specific IT Glue org IDs to exclude (comma-separated, optional): ").strip()
    exclude_ids = [e.strip() for e in exclude_ids_input.split(",") if e.strip()] if exclude_ids_input else []

    at_base_url = "https://webservices15.autotask.net/atservicesrest/v1.0"
    at_headers = {
        "Content-Type": "application/json",
        "ApiIntegrationCode": integration_code,
        "UserName": user_name,
        "Secret": secret
    }

    syncing_orgs = get_autotask_syncing_orgs(api_key)
    print(f"\n[INFO] Active Autotask orgs: {len(syncing_orgs)}")

    if exclude_ids:
        print(f"[INFO] Excluding {len(exclude_ids)} org(s): {', '.join(exclude_ids)}")

    contacts_to_create = []

    for org_id, name, included in syncing_orgs:
        # Skip excluded orgs
        if org_id in exclude_ids:
            print(f"[SKIP] Org '{name}' (ID: {org_id}) is in exclude list. Skipping...")
            continue

        autotask_remote_id = get_autotask_remote_id_from_included(included)
        if not autotask_remote_id:
            print(f"[SKIP] Org '{name}' (ID: {org_id}) - no Autotask ID found.")
            continue

        print(f"[INFO] Org: {name} (ID: {org_id}) - Autotask ID: {autotask_remote_id}")
        
        # Pass the licensed/unlicensed filter
        ms_only_contacts = get_microsoft_only_contacts(api_key, org_id, licensed)

        for _, contact_data in ms_only_contacts:
            attrs = contact_data.get("data", {}).get("attributes", {})
            emails, phones = extract_emails_and_phones(contact_data)
            if emails:
                contacts_to_create.append((autotask_remote_id, {
                    "firstName": attrs.get("first-name", ""),
                    "lastName": attrs.get("last-name", ""),
                    "email": emails[0],
                    "phone": phones[0] if phones else ""
                }))

    print(f"\n[INFO] Total contacts to create (before duplicates filtered): {len(contacts_to_create)}")
    if input("Press Enter to proceed, or any other key to cancel: ").strip() != "":
        sys.exit("Cancelled by user.")

    existing_contacts_cache = {}
    created_count = 0
    skipped = []

    with tqdm(total=len(contacts_to_create), desc="Creating contacts") as pbar:
        for company_id, contact_payload in contacts_to_create:
            if company_id not in existing_contacts_cache:
                existing_contacts_cache[company_id] = get_existing_autotask_contacts(at_base_url, at_headers, company_id)

            success = create_contact_in_autotask(at_base_url, at_headers, company_id, contact_payload, existing_contacts_cache[company_id])
            if success:
                created_count += 1
            else:
                skipped.append((contact_payload["firstName"], contact_payload["lastName"]))
            pbar.update(1)

    print(f"\n[RESULT] Created {created_count}/{len(contacts_to_create)} contacts.")
    if skipped:
        print(f"[SKIPPED] {len(skipped)} skipped or failed:")
        for f, l in skipped:
            print(f" - {f} {l}")
    else:
        print("[SKIPPED] 0 contacts failed.")


if __name__ == "__main__":
    main()

