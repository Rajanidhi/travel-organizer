# Run automatically by GitHub Actions: reads new emails and saves reservations.
import core

result = core.check_inbox(max_emails=10)
if not result["enabled"]:
    print("Email checking is not configured (EMAIL_ADDRESS / EMAIL_APP_PASSWORD missing).")
else:
    processed = result["processed"]
    saved = sum(1 for p in processed if p["status"].startswith("saved"))
    errors = sum(1 for p in processed if p["status"].startswith("error"))
    # Only counts are printed (no email subjects), so nothing private appears in public logs.
    print(f"Processed {len(processed)} new email(s): {saved} with reservations, {errors} errors.")
