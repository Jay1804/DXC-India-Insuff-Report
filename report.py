"""Data access + summary logic for the DXC India Insuff report."""
import os
from datetime import date, datetime

import pymysql
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

# Stage filters applied in the template's PivotTables:
#   Less Then Four Follow-Up -> followup-1 .. followup-4
#   More Then Four Follow-Up -> greater than followup-4
FOLLOWUP_THRESHOLD = 4

_FLEX = ",\n  ".join(f"CaseField.CASE_FLEX_FIELD{i}" for i in range(1, 18))

# "Candidate Insuff Calling Details New Logic" - trimmed to the columns used in the DXC India Insuff report.
# %% is a literal % (pymysql treats % as a parameter marker).
REPORT_SQL = f"""
SELECT
  InsuffCallingFollowup.CASE_CHECK_ID,
  CaseMaster.CASE_ARS_NO,
  IF(ec.client_external_id=2548,'TATA Motors',Company.company_name) AS company_name,
  CONCAT_WS(' ',Candidate.FIRST_NAME,Candidate.MIDDLE_NAME,Candidate.LAST_NAME) AS CANDIDATE_NAME,
  PROCESS1.PROCESS_NAME,
  CompanyLocation.OFFICE_NAME AS location,
  CaseMaster.RECEIVED_DATE,
  CaseInsuffDetail.INSUFF_RAISED_DATE,
  CaseChecks.insuff_date,
  (SELECT ecid.INSUFFICIENCY FROM ec_case_insuff_details ecid
     WHERE InsuffCallingFollowup.CASE_CHECK_ID=ecid.case_check_id
     ORDER BY ecid.insuff_id DESC LIMIT 1) AS insuff_remarks,
  checkpoint_live.fn_case_status(case_status) AS Case_status,
  checkpoint_live.fn_check_status(check_status) AS check_status,
  CaseChecks.CHECK_NAME,
  IF(CaseChecks.family_id=5,emc.company_name,emei.institute_name) AS verification_source,
  {_FLEX},
  Candidate.MOBILE_NO AS `Candidate contact no`,
  Candidate.EMAIL_ADDRESS AS `Candidate email id`,
  CaseChecks.reopen_date AS check_reopen_date,
  CONCAT('followup','-',InsuffCallingFollowup.stage) AS stage,
  InsuffCallingFollowup.ADDITIONAL_REMARKS,
  InsuffCallingFollowup.FOLLOWUP_DATE,
  InsuffCallingFollowup.NEXT_FOLLOWUP_DATE,
  (CASE InsuffCallingFollowup.REMARKS
     WHEN 1 THEN 'Does not have any details'
     WHEN 2 THEN 'Candidate has left the job'
     WHEN 3 THEN 'Not Reachable/connection not established'
     WHEN 4 THEN 'No Response'
     WHEN 5 THEN 'Candidate responded and will share details /documents'
     WHEN 6 THEN 'All follow-ups done, no response'
     WHEN 7 THEN 'Candidate will provide at a later stage'
     WHEN 8 THEN 'Refused to provide Details/Documents'
     WHEN 9 THEN 'Do Not Call'
     WHEN 10 THEN 'Contact details not provided'
     WHEN 11 THEN 'Document/Information received and sent for mapping'
     WHEN 12 THEN 'Approval Required From Client'
     WHEN 13 THEN 'Email sent but number not provided /Invalid Number'
     WHEN 14 THEN 'Email id not provided'
     WHEN 15 THEN 'Email sent & Calls follow ups done by IVR'
   END) AS remark,
  CaseChecks.reopen_date AS check_reopendate
FROM ec_follow_up_remarks InsuffCallingFollowup
INNER JOIN (
  SELECT ID, CASE_ID, USER_ID, CASE_CHECK_ID, IS_ACTIVE, CREATED, MAX(ID) AS latest
  FROM ec_follow_up_remarks
  WHERE IS_ACTIVE IN (0,1) AND CREATED >= %s AND CREATED <= %s AND CASE_CHECK_ID IS NOT NULL
  GROUP BY CASE_CHECK_ID
) r ON InsuffCallingFollowup.ID = r.latest AND InsuffCallingFollowup.CASE_CHECK_ID = r.CASE_CHECK_ID
INNER JOIN ec_insuff_calling_case InsuffCallingCase ON InsuffCallingCase.CASE_ID = InsuffCallingFollowup.CASE_ID
LEFT JOIN ec_case_insuff_details CaseInsuffDetail ON CaseInsuffDetail.CASE_ID = InsuffCallingCase.CASE_ID
INNER JOIN (SELECT CASE_ID, MAX(INSUFF_ID) AS latest1 FROM ec_case_insuff_details GROUP BY CASE_ID) s
  ON CaseInsuffDetail.INSUFF_ID = s.latest1 AND CaseInsuffDetail.CASE_ID = s.CASE_ID
LEFT JOIN ec_case_master CaseMaster ON CaseMaster.case_id = InsuffCallingFollowup.case_id
LEFT JOIN ec_client ec ON CaseMaster.client_id = ec.client_id
LEFT JOIN ec_case_candidates Candidate ON Candidate.candidate_id = CaseMaster.candidate_id
LEFT JOIN ec_case_fields CaseField ON CaseField.CASE_ID = CaseMaster.CASE_ID
LEFT JOIN ec_master_company Company ON InsuffCallingFollowup.CLIENT_ID = Company.COMPANY_ID
LEFT JOIN ec_client_process PROCESS1 ON CaseMaster.PROCESS_ID = PROCESS1.PROCESS_ID
LEFT JOIN ec_master_company_locations CompanyLocation ON CaseMaster.client_office_id = CompanyLocation.office_id
LEFT JOIN ec_case_checks CaseChecks ON InsuffCallingFollowup.CASE_CHECK_ID = CaseChecks.CASE_CHECK_ID
LEFT JOIN ec_case_check_verification_source eccvs ON InsuffCallingFollowup.case_check_id = eccvs.case_check_id
LEFT JOIN ec_master_company emc ON eccvs.org_id = emc.company_id
LEFT JOIN ec_master_educational_institute emei ON eccvs.org_id = emei.institute_id
WHERE InsuffCallingFollowup.IS_ACTIVE IN (0,1)
  AND InsuffCallingFollowup.CREATED >= %s
  AND InsuffCallingFollowup.CREATED <= %s
  AND Candidate.first_name NOT LIKE 'dummy%%'
  AND Candidate.first_name NOT LIKE 'test%%'
  AND client_external_id = 2307
  AND CaseChecks.check_status = 3 /* 3 = Insufficient */
  AND InsuffCallingFollowup.CASE_CHECK_ID IS NOT NULL
GROUP BY InsuffCallingFollowup.CASE_CHECK_ID
ORDER BY InsuffCallingFollowup.IS_ACTIVE
"""


def _connect():
    return pymysql.connect(
        host=os.environ["DB_HOST"],
        port=int(os.environ.get("DB_PORT", 3306)),
        database=os.environ["DB_NAME"],
        user=os.environ["DB_USER"],
        password=os.environ["DB_PASSWORD"],
        cursorclass=pymysql.cursors.DictCursor,
    )


def _plain(v):
    """Dates -> 'YYYY-MM-DD HH:MM:SS' strings (no timezone shifting); everything else untouched."""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    return v


def stage_no(stage):
    try:
        return int(str(stage).split("-", 1)[1])
    except (IndexError, ValueError):
        return None


def run_report(start: date, end: date) -> list[dict]:
    s, e = f"{start} 00:00:00", f"{end} 23:59:59"
    with _connect() as conn, conn.cursor() as cur:
        cur.execute(REPORT_SQL, (s, e, s, e))
        rows = [{k: _plain(v) for k, v in r.items()} for r in cur.fetchall()]
    for r in rows:
        r["followup_no"] = stage_no(r.get("stage"))  # same as the xlsx helper column AO
    return rows


def default_stages(rows):
    allst = sorted({r["followup_no"] for r in rows if r["followup_no"] is not None})
    return allst, [n for n in allst if n <= FOLLOWUP_THRESHOLD], [n for n in allst if n > FOLLOWUP_THRESHOLD]


def pivot(rows) -> list[tuple[str, int]]:
    """Excel-style pivot: Row Labels = remark (A-Z), values = count of remark. (blank) is hidden."""
    counts: dict[str, int] = {}
    for r in rows:
        if r.get("remark") is not None:
            counts[r["remark"]] = counts.get(r["remark"], 0) + 1
    return sorted(counts.items(), key=lambda kv: kv[0].lower())
