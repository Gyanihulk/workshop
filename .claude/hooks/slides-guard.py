#!/usr/bin/env python3
"""PreToolUse guard for Google Slides edits.

Normal edits pass silently (the allow rule in settings.json covers them). A batch that deletes
something which is not clearly a shape/image inside a slide is sent to the user for approval,
because deleting a whole slide is the one change that is hard to notice and easy to regret.

Element ids in this deck look like p3_i5, p8_g15, n_exp_07 or SLIDES_API123_4 (a suffix of
digits, optionally with i or g). Anything else that is being deleted is treated as a slide.
Ids it does not recognise therefore ask rather than allow.
"""
import json
import re
import sys

ELEMENT_ID = re.compile(r"(_[ig]?\d+$)|^SLIDES_API")


def main() -> None:
    try:
        event = json.load(sys.stdin)
    except Exception:
        return  # not our input: stay out of the way
    requests = (event.get("tool_input") or {}).get("requests") or []
    slides = sorted(
        {
            r["deleteObject"]["objectId"]
            for r in requests
            if isinstance(r, dict) and "deleteObject" in r and not ELEMENT_ID.search(r["deleteObject"].get("objectId", ""))
        }
    )
    if not slides:
        return
    json.dump(
        {
            "hookSpecificOutput": {
                "hookEventName": "PreToolUse",
                "permissionDecision": "ask",
                "permissionDecisionReason": "This edit deletes a whole slide: " + ", ".join(slides),
            }
        },
        sys.stdout,
    )


if __name__ == "__main__":
    main()
