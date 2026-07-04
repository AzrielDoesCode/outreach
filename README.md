zoho_outreach_sender.py

 Reads Asian_African_Financial_Institutions_v4_Researched.xlsx, renders the
 correct template (A/B/C) per row using the {{Placeholder}} system from
 Xcigence_Outreach_Master_Templates.md, and sends via Zoho SMTP.

 SAFETY DESIGN — read before running:
 1. DRY_RUN defaults to True. It will NOT send anything until you explicitly
    set DRY_RUN = False. First runs should always be dry.
 2. Rendered emails are written to /preview/ as .txt files so you can read
    every single one before any send happens. With 79 government regulators,
    exchanges, and credit bureaus on this list, a bad merge (wrong country's
    regulation, broken placeholder) going out is not something you can recall.
 3. BATCH_LIMIT caps how many emails one script run will actually send, and
    SEND_DELAY_RANGE adds a random pause between sends. Zoho's own usage
    policy states "Bulk/burst sending: Not supported" — sending 79 in a rapid
    loop is exactly the pattern that gets an account rate-limited or flagged,
    even though every email here is genuinely 1:1, not a marketing blast.
 4. Already-sent rows are tracked in the xlsx itself (a "Send Status" and
    "Send Timestamp" column are added) so re-running the script never
    double-sends to someone.
 5. Template D (Development Finance Institutions) and any UNASSIGNED rows
    are always skipped — those templates were never finalized/approved.

 Install: pip install openpyxl python-dotenv
