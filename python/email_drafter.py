"""
email_drafter.py
Stage 5: Automated Email Drafting
Generates a templated email to the LP with summary statistics
and a reference to the attached report.
"""

import sqlite3
import os
import logging
from datetime import datetime
from config import DB_PATH, EMAIL_TEMPLATE, LP_NAME, REPORT_PERIOD

logger = logging.getLogger(__name__)


def draft_email(db_path=None, output_dir="output"):
    """Generate the notification email for the LP."""
    if db_path is None:
        db_path = DB_PATH

    os.makedirs(output_dir, exist_ok=True)
    conn = sqlite3.connect(db_path)

    # Fetch summary statistics
    stats = conn.execute("""
        SELECT
            COUNT(DISTINCT fund_name) as fund_count,
            COUNT(DISTINCT gp_name) as gp_count,
            COUNT(*) as investment_count,
            COALESCE(SUM(commitment_eur), 0) as total_commitment,
            COALESCE(SUM(called_eur), 0) as total_called,
            COALESCE(SUM(distributed_eur), 0) as total_distributed
        FROM investments
    """).fetchone()

    fund_count, gp_count, inv_count, total_commit, total_called, total_dist = stats
    call_rate = total_called / total_commit if total_commit > 0 else 0
    dpi = total_dist / total_called if total_called > 0 else 0

    # Fetch validation stats
    issue_stats = conn.execute("""
        SELECT
            COUNT(*) as total,
            SUM(CASE WHEN severity = 'Critical' THEN 1 ELSE 0 END) as critical
        FROM validation_issues
    """).fetchone()

    issue_count, critical_count = issue_stats

    # Determine quality note
    if critical_count > 0:
        quality_note = (
            f"Please note: {critical_count} critical data quality issue(s) were identified "
            f"during processing. These have been flagged in the attached report under the "
            f"'Validation Issues' tab for your review. We recommend addressing these with "
            f"the relevant GP(s) before the next reporting cycle."
        )
    else:
        quality_note = "All records passed critical validation checks."

    conn.close()

    # Format the email
    email_body = EMAIL_TEMPLATE.format(
        lp_name=LP_NAME,
        fund_count=fund_count,
        investment_count=inv_count,
        gp_count=gp_count,
        total_commitment=total_commit,
        total_called=total_called,
        total_distributed=total_dist,
        call_rate=call_rate,
        dpi=dpi,
        total_records=inv_count,
        issue_count=issue_count,
        critical_count=critical_count,
        quality_note=quality_note,
    )

    # Save as text file
    email_path = os.path.join(output_dir, f"LP_Email_{REPORT_PERIOD.replace(' ', '_')}.txt")
    with open(email_path, "w") as f:
        f.write(email_body)

    # Also save as .eml format for email client import
    eml_path = os.path.join(output_dir, f"LP_Email_{REPORT_PERIOD.replace(' ', '_')}.eml")
    eml_content = build_eml(email_body, LP_NAME)
    with open(eml_path, "w") as f:
        f.write(eml_content)

    logger.info(f"Email drafted and saved to {email_path}")
    logger.info(f"EML version saved to {eml_path}")

    return email_path, eml_path


def build_eml(body, recipient_name):
    """Build a basic .eml file content."""
    now = datetime.now().strftime("%a, %d %b %Y %H:%M:%S +0100")

    eml = f"""From: pds-reporting@blackrock.com
To: {recipient_name.lower().replace(" ", ".")}@pensioenfonds.nl
Date: {now}
MIME-Version: 1.0
Content-Type: text/plain; charset="utf-8"
{body}"""

    return eml
