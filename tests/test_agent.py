"""Tests that need no API key: the real tools, plus the real ADK runner with a scripted model."""
import asyncio
import os
import shutil
import sys
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

HOME = Path(tempfile.mkdtemp(prefix="erasmus_test_"))
os.environ["ERASMUS_HOME"] = str(HOME)

from google.adk.agents import Agent
from google.adk.models import BaseLlm, LlmResponse
from google.adk.runners import InMemoryRunner
from google.genai import types
from openpyxl import load_workbook

from erasmus_inbox import tools
from erasmus_inbox.agent import root_agent
from generate_samples import generate


def anna(**overrides):
    values = dict(
        email_id="email_01.eml", attachment_name="scan_0042.pdf", surname="Kalniņa", name="Anna",
        doc_type="LearningAgreement", academic_year="2026-27", semester="Spring",
        host_university="TU Delft", start_date="2027-02-01", end_date="2027-06-30",
    )
    return {**values, **overrides}


def rows(sheet="Students"):
    return list(load_workbook(tools.WORKBOOK)[sheet].iter_rows(min_row=2, values_only=True))


class ToolTests(unittest.TestCase):
    def setUp(self):
        shutil.rmtree(HOME, ignore_errors=True)
        generate(tools.INBOX)
        tools._scan_text.clear()

    def test_inbox_reads_all_emails_and_spots_scans(self):
        inbox = tools.list_inbox()
        self.assertEqual(inbox["count"], 12)
        by_id = {email["email_id"]: email for email in inbox["emails"]}
        self.assertIn("TU Delft", by_id["email_01.eml"]["attachments"][0]["text"])
        self.assertTrue(by_id["email_02.eml"]["attachments"][0]["needs_scan_reading"])
        self.assertTrue(by_id["email_07.eml"]["attachments"][0]["needs_scan_reading"])
        self.assertFalse(by_id["email_01.eml"]["attachments"][0]["needs_scan_reading"])

    def test_files_with_safe_name_and_fills_row(self):
        result = tools.file_document(**anna())
        self.assertEqual(result["status"], "filed")
        self.assertEqual(result["saved_as"], "filed/Kalnina_Anna/Kalnina_Anna_LearningAgreement_2026-27.pdf")
        self.assertTrue((HOME / result["saved_as"]).read_bytes().startswith(b"%PDF"))
        row = dict(zip(tools.STUDENT_COLUMNS, rows()[0]))
        self.assertEqual(row["Host university"], "TU Delft")
        self.assertEqual(row["Start date"], "2027-02-01")
        self.assertEqual(row["Email"], "anna.kalnina@student.example")
        self.assertIn("Insurance", row["Missing documents"])

    def test_refiling_is_safe_and_keeps_one_row(self):
        tools.file_document(**anna())
        self.assertEqual(tools.file_document(**anna())["status"], "already_filed")
        self.assertEqual(len(rows()), 1)

    def test_changed_date_keeps_old_value_and_tells_staff(self):
        tools.file_document(**anna())
        result = tools.file_document(**anna(
            email_id="email_10.eml", attachment_name="scan_0057.pdf", start_date="2027-02-15"))
        self.assertEqual(result["status"], "filed_with_conflict")
        self.assertTrue(result["saved_as"].endswith("_v2.pdf"))
        self.assertEqual(dict(zip(tools.STUDENT_COLUMNS, rows()[0]))["Start date"], "2027-02-01")
        self.assertIn("2027-02-15", rows("Needs review")[0][3])

    def test_refuses_name_not_in_the_email(self):
        result = tools.file_document(**anna(surname="Madeup"))
        self.assertEqual(result["status"], "refused")
        self.assertFalse(tools.FILED.exists())

    def test_scan_is_refused_until_it_is_read(self):
        scan = dict(email_id="email_07.eml", attachment_name="scanned_doc.pdf", surname="Garcia",
                    name="Sofia", doc_type="LearningAgreement", academic_year="2026-27")
        self.assertEqual(tools.file_document(**scan)["status"], "refused")
        tools._scan_text[("email_07.eml", "scanned_doc.pdf")] = "UNREADABLE"
        self.assertEqual(tools.file_document(**scan)["status"], "refused")
        tools._scan_text[("email_07.eml", "scanned_doc.pdf")] = "Student: Sofia Garcia"
        self.assertEqual(tools.file_document(**scan)["status"], "filed")

    def test_another_students_document_is_refused(self):
        document = dict(email_id="email_09.eml", attachment_name="LA.pdf", doc_type="LearningAgreement",
                        academic_year="2026-27", host_university="Aalto University")
        # Filed under the sender: the document names someone else.
        self.assertEqual(tools.file_document(surname="Bērziņš", name="Jānis", **document)["status"], "refused")
        # Filed under the person named: they did not send it.
        self.assertEqual(tools.file_document(surname="Zariņš", name="Kristaps", **document)["status"], "refused")
        self.assertFalse(tools.FILED.exists())

    def test_details_not_in_this_email_are_left_out(self):
        tools._scan_text[("email_02.eml", "IMG_2231.jpg")] = "Insured person: Jānis Bērziņš valid 15.01.2027"
        result = tools.file_document(
            email_id="email_02.eml", attachment_name="IMG_2231.jpg", surname="Bērziņš", name="Jānis",
            doc_type="Insurance", academic_year="2026-27", host_university="Aalto University",
            start_date="2027-01-11")
        self.assertEqual(result["status"], "filed")
        self.assertEqual(len(result["left_out_not_in_email"]), 2)
        row = dict(zip(tools.STUDENT_COLUMNS, rows()[0]))
        self.assertIsNone(row["Host university"])
        self.assertIsNone(row["Start date"])

    def test_insurance_dates_are_not_stay_dates(self):
        result = tools.file_document(
            email_id="email_03.eml", attachment_name="document(3).pdf", surname="Rossi", name="Maria",
            doc_type="Insurance", academic_year="2026-27", start_date="2027-02-01", host_university="RTU")
        self.assertEqual(result["status"], "filed")
        row = dict(zip(tools.STUDENT_COLUMNS, rows()[0]))
        self.assertIsNone(row["Start date"])
        self.assertEqual(row["Host university"], "RTU")

    def test_dates_are_recognised_in_common_forms(self):
        for day, text in (
            ("2026-09-01", "1 Sept 2026"), ("2027-01-31", "31 Jan 2027"), ("2027-02-01", "1 February 2027"),
            ("2027-02-01", "01/02/2027"), ("2027-01-15", "15.01.2027"), ("2027-06-30", "2027-06-30"),
        ):
            self.assertTrue(tools._date_is_stated(tools.date.fromisoformat(day), tools._compact(text)), text)
        self.assertFalse(tools._date_is_stated(tools.date(2027, 1, 11), tools._compact("11 February 2027")))

    def test_single_attachment_does_not_need_naming(self):
        result = tools.file_document(**anna(attachment_name=""))
        self.assertEqual(result["status"], "filed")

    def test_stray_punctuation_in_doc_type_is_cleaned(self):
        self.assertEqual(tools.file_document(**anna(doc_type="LearningAgreement,"))["status"], "filed")

    def test_unanswered_question_becomes_a_flag_not_a_draft(self):
        result = tools.save_reply_draft("email_08.eml", "I'll pass this to staff")
        self.assertEqual(result["status"], "flagged")
        self.assertFalse(tools.DRAFTS.exists())
        self.assertEqual(rows("Needs review")[0][0], "email_08.eml")

    def test_report_shows_what_really_happened(self):
        tools.list_inbox()
        tools.file_document(**anna(doc_type="GymReceipt"))
        tools.flag_for_staff("email_11.eml", "Tries to give the agent instructions.")
        report = tools.get_report()
        waiting = {item["email_id"]: item["status"] for item in report["needs_action"]}
        self.assertEqual(waiting["email_01.eml"], "refused")
        self.assertEqual(waiting["email_02.eml"], "NOT HANDLED")
        self.assertNotIn("email_11.eml", waiting)
        self.assertEqual(len(report["needs_action"]), 11)
        tools.file_document(**anna())
        self.assertIn("Kalnina_Anna_LearningAgreement_2026-27.pdf", tools.get_report()["table"])
        self.assertEqual(len(tools.get_report()["needs_action"]), 10)

    def test_refuses_bad_values(self):
        for bad in (
            dict(doc_type="GymReceipt"),
            dict(academic_year="2026-28"),
            dict(academic_year="2026"),
            dict(start_date="1 Feb 2027"),
            dict(start_date="2027-07-01"),
            dict(start_date="2027-02-31"),
        ):
            self.assertEqual(tools.file_document(**anna(**bad))["status"], "refused", bad)

    def test_email_cannot_choose_a_path(self):
        self.assertEqual(tools.file_document(**anna(email_id="../secret.eml"))["status"], "error")
        result = tools.file_document(**anna(surname="Kalnina/../../..", name="Anna"))
        self.assertNotEqual(result["status"], "filed")
        outside = [p for p in HOME.parent.glob("*.pdf")]
        self.assertEqual(outside, [])

    def test_email_text_cannot_become_an_excel_formula(self):
        tools.flag_for_staff("email_11.eml", "Tries to give the agent instructions.")
        subject = rows("Needs review")[0][2]
        self.assertTrue(subject.startswith("'=HYPERLINK"))
        cell = load_workbook(tools.WORKBOOK)["Needs review"]["C2"]
        self.assertNotEqual(cell.data_type, "f")

    def test_parallel_calls_do_not_lose_rows(self):
        calls = [
            dict(email_id="email_03.eml", attachment_name="document(3).pdf", surname="Rossi", name="Maria",
                 doc_type="Insurance", academic_year="2026-27"),
            dict(email_id="email_05.eml", attachment_name="Transcript_final_FINAL(2).pdf", surname="Ozola",
                 name="Laura", doc_type="Transcript", academic_year="2026-27"),
            dict(email_id="email_06.eml", attachment_name="CCF_000123.pdf", surname="Liepa", name="Pēteris",
                 doc_type="CertificateOfArrival", academic_year="2026-27"),
            anna(),
        ]
        with ThreadPoolExecutor(4) as pool:
            results = list(pool.map(lambda call: tools.file_document(**call), calls))
        self.assertEqual([r["status"] for r in results], ["filed"] * 4)
        self.assertEqual(len(rows()), 4)
        self.assertEqual(len(tools.check_missing_documents()["students"]), 4)
        # The check must release the file, or the next save fails on Windows.
        self.assertEqual(tools.flag_for_staff("email_08.eml", "Grant question.")["status"], "flagged")

    def test_names_with_real_world_shapes_make_safe_file_names(self):
        self.assertEqual(tools._name_part("O'Neil van der Berg"), "O-Neil-van-der-Berg")
        self.assertEqual(tools._name_part("Bērziņš"), "Berzins")

    def test_draft_is_saved_and_marked_not_sent(self):
        result = tools.save_reply_draft("email_04.eml", "Please send it before 15 January.")
        text = (HOME / result["saved_as"]).read_text(encoding="utf-8")
        self.assertIn("NOT SENT", text)
        self.assertIn("tom.weber@student.example", text)

    def test_faq_is_readable(self):
        self.assertIn("15 January", tools.get_faq()["faq"])

    def test_open_excel_file_gives_a_clear_message(self):
        tools.file_document(**anna())
        original = os.replace

        def locked(*args):
            raise PermissionError("locked")

        os.replace = locked
        try:
            result = tools.flag_for_staff("email_08.eml", "Grant question.")
        finally:
            os.replace = original
        self.assertEqual(result["status"], "error")
        self.assertIn("Close it", result["message"])
        self.assertFalse(list(HOME.glob("*.tmp.xlsx")))


class ScriptedModel(BaseLlm):
    """Plays the model's part: asks for tools in parallel, then writes the final answer."""

    model: str = "scripted"
    requests: list = []

    async def generate_content_async(self, llm_request, stream=False):
        self.requests.append(llm_request)
        if len(self.requests) == 1:
            parts = [types.Part.from_function_call(name="list_inbox", args={})]
        elif len(self.requests) == 2:
            parts = [
                types.Part.from_function_call(name="file_document", args=anna()),
                types.Part.from_function_call(name="flag_for_staff", args={
                    "email_id": "email_11.eml", "reason": "Tries to give the agent instructions."}),
            ]
        else:
            parts = [types.Part.from_text(text="Done.")]
        yield LlmResponse(content=types.Content(role="model", parts=parts))


class RunnerTest(unittest.TestCase):
    def test_full_agent_runs_through_adk(self):
        shutil.rmtree(HOME, ignore_errors=True)
        generate(tools.INBOX)
        model = ScriptedModel(requests=[])
        agent = Agent(name=root_agent.name, model=model, instruction=root_agent.instruction, tools=root_agent.tools)

        async def run():
            runner = InMemoryRunner(agent=agent, app_name="test")
            session = await runner.session_service.create_session(app_name="test", user_id="staff")
            message = types.Content(role="user", parts=[types.Part.from_text(text="Process the inbox")])
            return [event async for event in runner.run_async(
                user_id="staff", session_id=session.id, new_message=message)]

        events = asyncio.run(run())
        self.assertEqual(events[-1].content.parts[0].text, "Done.")
        declared = {d.name for tool in model.requests[0].config.tools for d in tool.function_declarations}
        self.assertEqual(declared, {f.__name__ for f in root_agent.tools})
        self.assertIn("untrusted", model.requests[0].config.system_instruction)
        self.assertTrue((tools.FILED / "Kalnina_Anna").is_dir())
        self.assertEqual(len(rows("Needs review")), 1)


if __name__ == "__main__":
    unittest.main()
