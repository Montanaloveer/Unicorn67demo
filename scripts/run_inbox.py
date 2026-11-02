"""Run the agent on the inbox from the terminal and log every tool call. A backup for adk web."""
import asyncio
import json
import os
import sys
import time
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
load_dotenv(ROOT / "erasmus_inbox" / ".env")

from google.adk.runners import InMemoryRunner
from google.genai import types

from erasmus_inbox.agent import root_agent


def short(value, limit=300):
    text = json.dumps(value, ensure_ascii=False)
    return text if len(text) <= limit else text[:limit] + "..."


async def main(prompt):
    runner = InMemoryRunner(agent=root_agent, app_name="erasmus_inbox")
    session = await runner.session_service.create_session(app_name="erasmus_inbox", user_id="staff")
    message = types.Content(role="user", parts=[types.Part.from_text(text=prompt)])
    started, model_turns = time.time(), 0
    async for event in runner.run_async(user_id="staff", session_id=session.id, new_message=message):
        if not event.content or not event.content.parts:
            continue
        if event.content.role == "model":
            model_turns += 1
        for part in event.content.parts:
            if part.function_call:
                print(f"CALL   {part.function_call.name} {short(dict(part.function_call.args or {}))}")
            elif part.function_response:
                response = part.function_response.response
                if part.function_response.name == "list_inbox":
                    response = {"count": response.get("count")}
                print(f"RESULT {part.function_response.name} {short(response)}")
            elif part.text:
                print(f"\n{part.text}\n")
    print(f"--- {model_turns} model turns in {time.time() - started:.0f}s")


if __name__ == "__main__":
    try:
        asyncio.run(main(" ".join(sys.argv[1:]) or "Process the inbox"))
    except Exception as error:
        # Never print the key, even if a library puts it in an error message.
        print("FAILED:", str(error).replace(os.environ.get("GOOGLE_API_KEY") or "\0", "[key hidden]"))
        sys.exit(1)
