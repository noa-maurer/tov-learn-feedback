"""
feedback_tool.py — כלי לניתוח משובים על Tov-Learn

שימוש:
  python feedback_tool.py validate --index 0       # ולידציה לשורה 0 (מ-feedbacks.csv)
  python feedback_tool.py validate --all            # ולידציה לכל השורות (מ-feedbacks.csv)
  python feedback_tool.py analyze --index 0         # ניתוח לשורה 0 (מ-feedbacks.csv)
  python feedback_tool.py analyze --all             # ניתוח לכל השורות (מ-feedbacks.csv)
  python feedback_tool.py sync                      # סנכרון מ-Google Sheets → ניתוח → כתיבה חזרה
  python feedback_tool.py sync --list-columns       # הצגת שמות העמודות בלבד (לאבחון)
"""

import os
import sys
import json
import re
import time
import argparse
from collections import defaultdict
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from google import genai
from google.genai import types
from dotenv import load_dotenv

# Windows consoles often default stdout/stderr to a non-UTF-8 codepage (e.g. cp1255),
# which can't encode the Hebrew text and emoji this script prints.
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

# Resolve local files relative to this script, so the folder works on any machine/user.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

load_dotenv(os.path.join(BASE_DIR, ".env"))

# ─── Gemini ───────────────────────────────────────────────────────────────────
MODEL = "gemini-3.5-flash-lite"
client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))

# ─── Google Sheets ────────────────────────────────────────────────────────────
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "service_account.json")
SPREADSHEET_ID       = "14mdsbX_aHo29xJhcsGU52lcw5tOvhb7i2anwA5dXdEY"
RESPONSES_GID        = 2076837780   # Tab 1 — Google Form responses
ANALYSIS_GID         = 2047537741   # Tab 2 — analyzed output
SYNC_STATE_FILE      = os.path.join(BASE_DIR, "sync_state.json")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Candidates for the email column name in Sheet 1.
# If your column name is different, add it here.
EMAIL_COL_CANDIDATES = ["Email Address", "כתובת דוא\"ל", "אימייל", "email", "Email", "כתובת אימייל", "כתובת אימייל:"]
NAME_COL_CANDIDATES  = ["שם", "Name", "Full Name", "שם מלא"]

# ─── Questions ────────────────────────────────────────────────────────────────
QUESTIONS = [
    "מה עבד טוב?",
    "מה היה לא ברור או קשה?",
    "מה חשוב לתקן/לשפר? מה היה חסר?",
    "אילו חלקים היו שימושיים במיוחד?",
    "האם יש לכם רעיונות חדשים להוסיף?"
]

# ─── System prompts ───────────────────────────────────────────────────────────
VALIDATOR_SYSTEM = """אתה מנתח משוב על מערכת לימוד מבוססת AI בשם Tov-Learn.
המערכת מבוססת על Claude Code ומאפשרת ללמוד חומר תיאורטי, לשאול שאלות, לפתור תרגילים ולהיבחן — הכל בשיחה עם ה-AI.

תפקידך: לבדוק אם תשובה במשוב מספיקה לניתוח. תשובה מספיקה היא תשובה שמכילה טענה — כלומר, אומרת משהו ברור, גם אם זה "כלום".

תשובה מספקת היא אחת מהבאות:
- תשובת "אין" — המשתתף מציין במפורש שאין לו קשיים, רעיונות, תלונות וכו' (למשל: "לא היו לי קשיים", "הכל היה בסדר", "אין לי רעיונות להוסיף"). זו תשובה תקפה — היא עונה על השאלה.
- מספר טענות ספציפיות, גם בלי דוגמה נרטיבית (מספיק לציין מה, מדוע, או כיצד)
- פיצ'ר, חלק, שיעור, או היבט ספציפי של המערכת
- הסבר ברור מדוע משהו עבד, לא עבד, היה חסר, או היה שימושי

דוגמאות לתשובות מספקות:
- "הכל היה בסדר." / "לא נתקלתי בקשיים." / "אין לי רעיונות להוסיף." (תשובת "אין" — ענתה על השאלה)
- "הבחן בסוף שיעור 2.4 עזר לי לגלות שלא הבנתי את ensure_ascii=False — זה גרם לי לחזור ולתרגל שוב."
- "קצב וצורת לימוד מותאמים לרמה. אפשרות לדון עם המודל כדי להבין את החומר. בדיקה שהחומר הובן. התנסות בתרגילים ושיעורי בית."

תשובה אינה מספקת רק כשהיא מעלה טענה (חיובית או שלילית) אבל לא אומרת שום דבר ספציפי לגביה:
- "הבחנים עזרו לי לדעת כמה אני שולטת בחומר." (טענה ללא פירוט — לא ברור מה בדיוק עזר)
- "השיעור היה מבלבל." (טענה ללא פירוט — לא ברור איזה שיעור, מה היה מבלבל)
- "לא הבנתי את התרגילים." (טענה ללא פירוט — לא ברור אילו תרגילים, מה לא היה ברור)

אם התשובה אינה מספקת, כתוב שאלת המשך קצרה וספציפית בעברית שתעזור למשתתף לפרט.
החזר JSON בלבד, ללא טקסט נוסף: {"sufficient": true/false, "followup": "שאלת המשך" או null}"""

ANALYZER_SYSTEM = """אתה מנתח משוב על מערכת לימוד מבוססת AI בשם Tov-Learn.
קיבלת תשובה של אחד המשתתפים לשאלה אחת. תפקידך לסכם את הנקודות המרכזיות לשאלה זו.

כללים בסיסיים:
- הישאר צמוד למה שנכתב — אל תמציא מידע ואל תשנה את משמעות הטענות.
- סכם בעברית יומיומית ופשוטה — אל תשתמש במילים פורמליות או אקדמיות.
- כתוב בעברית בלבד — אסור להשתמש בתווים ערביים.
- כאשר הטקסט המקורי כולל מונח, ביטוי, או שם בשפה אחרת (כמו "compact", "Vibe Coding", "Lesson Summaries", "Lesson") — שמור על הכתיב המקורי. אל תתעתק אותו לעברית. לדוגמה: "compact" נשאר "compact" ולא "קומפקט"; "Vibe Coding" נשאר "Vibe Coding" ולא "ויב קודינג".
- אם תשובה ריקה, כתוב "לא ענה".
- אל תשתמש במילה "חובה" — השתמש ב"חשוב".

כלל הנקודות — קריטי:
ספור נקודות בלט (bullets) רק לפי מספר הטענות, לא לפי מספר המשפטים.
טענה אחת = נקודה אחת, גם אם היא כוללת סיבות, דוגמאות, פרטים, השוואות, או מרכיבים שונים.

כיצד לזהות שמדובר בטענה אחת:
- אם כל הפרטים מסבירים למה משהו עבד, לא עבד, חסר, או שימושי — זו טענה אחת.
- אם השוואה לעבודה אמיתית באה להסביר למה הבעיה קיימת — היא חלק מאותה טענה, לא טענה נפרדת.
- אם רעיון כולל מספר מרכיבים (מבנה קובץ, שפה, תוכן) שכולם שייכים לאותו פיצ'ר — זו טענה אחת.

מתי זה כן כמה טענות נפרדות:
אם הטקסט מפרט כמה בעיות או נקודות שונות ובלתי-קשורות זו לזו (נושאים שונים, סיבות שונות, היבטים שונים של המערכת), שרק מוזכרות ברצף באותו משפט או פסקה בלי שכולן מסבירות טענה משותפת אחת — אלה טענות נפרדות, נקודה לכל בעיה. אל תאחד בעיות שונות רק בגלל שהמשתתף כתב אותן ברצף אחד.

דוגמה נכונה — טענה אחת, נקודה אחת:
"כדאי שיהיה קובץ סיכומי שיעור שמתעדכן אחרי כל שיעור, עם כותרת ראשית, כותרת משנה לכל מודול, ותת-כותרת לכל שיעור עם הסיכום. אפשר לבקש את הסיכום בעברית, באנגלית, או בשני קבצים נפרדים — אחד לכל שפה."

דוגמה שגויה — אל תעשה כך:
- כדאי שיהיה קובץ סיכומי שיעור.
- כדאי שתהיה אפשרות לבחור שפה.
- כדאי שיהיה מבנה מסודר לקובץ.

דוגמה לכמה בעיות שונות שהוזכרו ברצף — כאן כן צריך לפצל:
"הלמידה העצמית מול המסך לא הייתה קלה. קובץ ה-script לא תורגם למצגת בדרך רגילה. הפלט בטרמינל הופיע בעברית הפוכה. העדפתי כלי אחד על פני אחר וזה לא תמיד התאפשר."
שגוי — נקודה אחת שמאחדת ארבע בעיות לא קשורות:
- הלמידה העצמית לא הייתה קלה, קובץ ה-script לא תורגם למצגת כרגיל, הפלט בטרמינל הופיע הפוך, וההעדפה לכלי מסוים לא תמיד התאפשרה.
נכון — ארבע נקודות נפרדות, אחת לכל בעיה:
- הלמידה העצמית מול המסך לא הייתה קלה.
- קובץ ה-script לא תורגם למצגת בדרך הרגילה.
- הפלט בטרמינל הופיע בעברית הפוכה.
- ההעדפה לעבוד עם כלי מסוים על פני אחר לא תמיד התאפשרה.

כיצד לנסח השוואה לעבודה אמיתית:
השוואה לעבודה אמיתית היא תמיד חלק מהטענה שהיא מסבירה — לעולם לא נקודה נפרדת בפני עצמה.
אם הטענה היא "ה-AI לא יודע מה התכוונו המנחים" וההשוואה היא "כמו שבעבודה אמיתית רק המנהל יודע" — זו נקודה אחת בלבד.
שגוי — שתי נקודות:
- ה-AI לא ידע לענות על שאלות שמחוץ להקשר.
- גם בפרויקטים אמיתיים, רק המנהל יודע מה עבר בראש שלו.
נכון — נקודה אחת:
- ה-AI לא ידע לענות על שאלות שמחוץ להקשר — בדיוק כמו שבפרויקטים אמיתיים רק המנהל או הלקוח שביקשו את הפרויקט יודעים מה עבר בראש שלהם.

שמירה על התנהגות של גורמים חיצוניים:
אם המשתתף מציין שגורם חיצוני (כמו מנחים, מנהל, לקוח) ביקש ממנו משהו, ציפה ממנו משהו, או הפנה אותו לכיוון מסוים — שמור על זה בסיכום. אל תמחק את הסוכן.
לדוגמה: אם המשתתף כתב "המנחים ציפו שאבין לבד או אשאל את ה-AI, במקום לענות לי ישירות", אל תסכם את זה כ"היה קשה לקבל תשובות" — הסכם כ"המנחים ציפו שהמשתתף יבין לבד או ישאל את ה-AI, אבל...".

הפרדה בין קושי אישי מול גורם אנושי לבין מגבלה מבנית של ה-AI:
אם המשתתף מתאר קושי אישי שלו (כמו קשיי תקשורת, נטייה לפרש לא נכון) שגרם לו קושי להבין דבר-מה לבד או מול גורם אנושי (כמו מנחים), וגם מגבלה מבנית נפרדת של ה-AI (כמו חוסר הקשר, אי-ידיעה של כוונות גורם חיצוני) — אלה שתי טענות נפרדות, גם אם מוזכרות ברצף אחד. אל תמזג אותן ואל תייחס את התכונה האישית של המשתתף ל-AI עצמו. ה-AI לא ידע לענות מסיבה מבנית (חוסר הקשר לכוונות הגורם החיצוני) — לא בגלל קשיי תקשורת של המשתתף.
שגוי — טענה אחת שמייחסת קושי תקשורת ל-AI:
- היה קשה לשאול את ה-AI שאלות שלא קשורות לחומר, כי ה-AI לא ידע מה המנחים התכוונו, בשל קשיי תקשורת ונטייה לפרש דברים לא נכון.
נכון — שתי טענות נפרדות:
- היה קשה להבין לבד הגדרות שקבעו המנחים, כי יש למשתתף קשיי תקשורת ונטייה לפרש דברים לא נכון.
- ה-AI לא ידע לענות על שאלות שמחוץ להקשר של החומר, כי הוא לא יכול לדעת מה המנחים התכוונו — בדיוק כמו שבפרויקטים אמיתיים רק המנהל או הלקוח יודעים מה עבר בראש שלהם.

שמירה על בעלות אישית של תכונות:
אם המשתתף מתאר תכונה אישית שלו (כמו "יש לי קשיי תקשורת" או "אני נוטה לפרש דברים לא נכון"), הצג אותה כתכונה אישית — לא כמצב מצבי כללי.
שגוי: "והיו קשיי תקשורת" / "היו קשיי תקשורת"
נכון: "ויש לה/לו קשיי תקשורת" / "ויש לה/לו נטייה לפרש דברים לא נכון"

אל תיצור ישויות שלא קיימות:
אם המשתתף מתאר "המנהל או הלקוח שביקשו את הפרויקט" — הם הם מי שביקש. אל תפצל לשתי ישויות נפרדות ("מי שביקש" + "המנהל"). כתוב: "המנהל או הלקוח שביקשו את הפרויקט יודעים מה עבר בראש שלהם".

בחירת מילים — בהתבסס על לעומת כולל:
כשהסיכום אמור להיות מבוסס על משהו (ולא להכיל אותו כפריט נוסף) — השתמש ב"בהתבסס על", לא ב"כולל".

שמירה על זמן תנאי/היפותטי:
כשהמשתתף משתמש בניסוח היפותטי או תנאי ("היה לוקח", "היה עוזר", "היה חוסך" — כלומר מה *היה* קורה אם לא היה X), שמור על אותה צורת לשון בסיכום. אל תמיר לעבר פשוט.
שגוי: "...שלקח הרבה זמן"
נכון: "...שהיה לוקח הרבה זמן" """


# ─── Core API helpers ─────────────────────────────────────────────────────────
def call_api(system_prompt: str, user_prompt: str) -> str:
    from google.genai.errors import ClientError
    for attempt in range(4):
        try:
            response = client.models.generate_content(
                model=MODEL,
                contents=user_prompt,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    temperature=0.3,
                )
            )
            time.sleep(4)   # proactive throttle — stays under 15 RPM free-tier limit
            return response.text
        except ClientError as e:
            if e.status_code == 429 and attempt < 3:
                # Parse "Please retry in X.Xs" from the error message
                m = re.search(r'retry in (\d+(?:\.\d+)?)', str(e))
                wait = float(m.group(1)) + 3 if m else 20
                print(f"  ⏳ rate limit — ממתין {wait:.0f} שניות ומנסה שוב...")
                time.sleep(wait)
            else:
                raise


# ─── Nothing-to-report detection ─────────────────────────────────────────────
_NOTHING_MARKERS = [
    "אין לי", "אין לנו", "לא היה", "לא היו", "לא היה לי",
    "הכל בסדר", "הכל היה בסדר", "הכל עבד", "לא נתקלתי",
    "לא מצאתי", "אין רעיונות", "לא עלה", "לא חסר", "לא קשה",
]

def _is_nothing_to_report(answer: str) -> bool:
    """Short answer whose entire content is 'nothing to report'."""
    a = answer.strip()
    if len(a) > 150:   # long answers always contain actual content
        return False
    return any(m in a for m in _NOTHING_MARKERS)


def parse_json(text: str) -> dict:
    match = re.search(r'\{.*?\}', text, re.DOTALL)
    if not match:
        return {"sufficient": True, "followup": None}
    try:
        return json.loads(match.group())
    except json.JSONDecodeError:
        return {"sufficient": True, "followup": None}


# ─── Local CSV commands (existing) ────────────────────────────────────────────
def validate_row(row_index: int, df: pd.DataFrame):
    row = df.loc[row_index]
    name = row.get("שם", f"משתתף {row_index}")
    print(f"\n{'='*50}")
    print(f"ולידציה — {name} (שורה {row_index})")
    print(f"{'='*50}\n")

    needs_followup = []

    for q in QUESTIONS:
        answer = str(row.get(q, "")).strip()
        if not answer or answer == "nan":
            print(f"⚠️  {q}\n    לא ענה על שאלה זו.\n")
            continue

        prompt = (
            f'שאלה: "{q}"\n'
            f'תשובה: "{answer}"\n\n'
            f'האם התשובה מכילה דוגמה קונקרטית או מצב ספציפי?\n'
            f'החזר JSON בלבד: {{"sufficient": true/false, "followup": "שאלת המשך או null"}}'
        )
        result = parse_json(call_api(VALIDATOR_SYSTEM, prompt))

        if result.get("sufficient"):
            print(f"✅ {q}\n    התשובה מכילה דוגמה קונקרטית.\n")
        else:
            followup = result.get("followup", "")
            print(f"❌ {q}\n    שאלת המשך לשליחה: {followup}\n")
            needs_followup.append((q, followup))

    if needs_followup:
        print(f"\n⚠️  נדרשות שאלות המשך ({len(needs_followup)} שאלות).")
        print("    מומלץ לשלוח את שאלות המשך לפני הניתוח.\n")
    else:
        print("\n✅ כל התשובות מספקות — ניתן להמשיך לשלב הניתוח.\n")


def analyze_row(row_index: int, df: pd.DataFrame):
    row = df.loc[row_index]
    name = row.get("שם", f"משתתף {row_index}")
    print(f"\n{'='*50}")
    print(f"ניתוח — {name} (שורה {row_index})")
    print(f"{'='*50}\n")

    answers_text = ""
    for q in QUESTIONS:
        answer = str(row.get(q, "")).strip()
        answers_text += f"שאלה: {q}\nתשובה: {answer if answer and answer != 'nan' else 'לא ענה'}\n\n"

    prompt = f"להלן משוב מאחד המשתתפים על Tov-Learn:\n\n{answers_text}\nסכם את הנקודות המרכזיות לפי כל שאלה."
    result = call_api(ANALYZER_SYSTEM, prompt)
    print(result)


# ─── Google Sheets helpers ────────────────────────────────────────────────────
def get_sheets():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    responses_ws = spreadsheet.get_worksheet_by_id(RESPONSES_GID)
    analysis_ws  = spreadsheet.get_worksheet_by_id(ANALYSIS_GID)
    return spreadsheet, responses_ws, analysis_ws


def _find_col(record: dict, candidates: list) -> str | None:
    for c in candidates:
        if c in record:
            return c
    return None


def read_responses(responses_ws) -> list[dict]:
    """Read Sheet 1, deduplicate by email (keep last row), return list of row dicts."""
    records = responses_ws.get_all_records()
    if not records:
        return []

    email_col = _find_col(records[0], EMAIL_COL_CANDIDATES)
    if not email_col:
        print("⚠️  לא נמצאה עמודת אימייל — לא ניתן לבצע deduplicate.")
        print(f"    עמודות קיימות: {list(records[0].keys())}")
        print("    הוסף את שם העמודה ל-EMAIL_COL_CANDIDATES ונסה שוב.")
        print("    ממשיך ללא deduplicate (כל שורה = משתתף נפרד).")
        return records

    seen: dict[str, dict] = {}
    for record in records:
        email = str(record.get(email_col, "")).strip()
        key = email if email else f"__nomail_{len(seen)}"
        seen[key] = record   # later submission overwrites earlier one

    return list(seen.values())


def read_local_feedback() -> list[tuple[int, dict]]:
    """Read feedbacks.csv (local/manual feedback, e.g. the tool author's own) as (index, row_dict) pairs."""
    if not os.path.exists("feedbacks.csv"):
        return []
    df = pd.read_csv("feedbacks.csv", index_col=0, encoding="utf-8-sig")
    return [(int(idx), row.to_dict()) for idx, row in df.iterrows()]


def load_sync_state() -> dict:
    if os.path.exists(SYNC_STATE_FILE):
        with open(SYNC_STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_sync_state(state: dict):
    with open(SYNC_STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def process_response(row_dict: dict) -> dict[str, str]:
    """Validate then analyze each question individually.
    Returns dict: question -> cell text (analyzed or flagged)."""
    results: dict[str, str] = {}

    for q in QUESTIONS:
        answer = str(row_dict.get(q, "")).strip()
        if not answer or answer == "nan":
            results[q] = "לא ענה"
            continue

        # Step 1 — validate
        val_prompt = (
            f'שאלה: "{q}"\n'
            f'תשובה: "{answer}"\n\n'
            f'האם התשובה מכילה דוגמה קונקרטית או מצב ספציפי?\n'
            f'החזר JSON בלבד: {{"sufficient": true/false, "followup": "שאלת המשך או null"}}'
        )
        val = parse_json(call_api(VALIDATOR_SYSTEM, val_prompt))

        if not val.get("sufficient"):
            followup = val.get("followup", "")
            results[q] = f"⚠️ תשובה לא מספיקה\nשאלת המשך: {followup}\n\n{answer}"
            print(f"  ❌ {q[:40]}...")
        elif _is_nothing_to_report(answer):
            # Valid "nothing to report" — readable in Sheet 2, filtered from aggregate
            results[q] = "— " + answer
            print(f"  — {q[:40]}... (אין נתונים לדיווח)")
        else:
            # Step 2 — analyze this question alone
            analyze_prompt = (
                f"להלן משוב מאחד המשתתפים על Tov-Learn:\n\n"
                f"שאלה: {q}\n"
                f"תשובה: {answer}\n\n"
                f"סכם את הנקודות המרכזיות לשאלה זו בנקודות בלט."
            )
            results[q] = call_api(ANALYZER_SYSTEM, analyze_prompt)
            print(f"  ✅ {q[:40]}...")

    return results


def find_analysis_row_num(analysis_ws, index: int) -> int | None:
    """Return 1-based row number in Sheet 2 where column A == index, or None."""
    col_a = analysis_ws.col_values(1)
    for i, val in enumerate(col_a):
        if str(val) == str(index):
            return i + 1
    return None


def write_analysis_row(analysis_ws, index: int, results: dict[str, str], update_row: int | None):
    """Append a new row, or update an existing one (preserving old values in (*...*)."""
    row_values = [str(index)] + [results.get(q, "") for q in QUESTIONS]

    if update_row is None:
        analysis_ws.append_row(row_values, value_input_option="USER_ENTERED")
        return

    # Update — merge with existing values
    existing = analysis_ws.row_values(update_row)
    while len(existing) < len(row_values):
        existing.append("")

    merged = []
    for new_val, old_val in zip(row_values, existing):
        # Preserve old valid analysis; don't preserve old flags (⚠️) or empty cells
        if old_val and old_val != new_val and not old_val.startswith("⚠️"):
            merged.append(f"{new_val}\n\n(*{old_val}*)")
        else:
            merged.append(new_val)

    end_col = chr(ord("A") + len(QUESTIONS))  # e.g. "F" for 5 questions
    analysis_ws.update(
        range_name=f"A{update_row}:{end_col}{update_row}",
        values=[merged],
        value_input_option="USER_ENTERED"
    )


# ─── Sync command ─────────────────────────────────────────────────────────────
def sync_command(list_columns: bool = False, force_index: int | None = None):
    print("מתחבר ל-Google Sheets...")
    _, responses_ws, analysis_ws = get_sheets()

    if list_columns:
        records = responses_ws.get_all_records()
        if records:
            print("\nעמודות ב-Sheet 1 (תגובות):")
            for col in records[0].keys():
                print(f"  • {col}")
        headers = analysis_ws.row_values(1)
        print("\nעמודות ב-Sheet 2 (ניתוח):")
        for col in headers:
            print(f"  • {col}")
        return

    state = load_sync_state()

    # ── Local feedback (e.g. the tool author's own, from feedbacks.csv) ──────
    for idx, row in read_local_feedback():
        state_key = f"_local_{idx}"
        current_answers = {q: str(row.get(q, "")).strip() for q in QUESTIONS}
        if force_index != idx and state.get(state_key, {}).get("answers") == current_answers:
            print(f"⏭️  [מקומי {idx}] — לא השתנה, מדלג")
            continue

        print(f"\n{'='*50}")
        print(f"מעבד [מקומי {idx}] (feedbacks.csv)")
        print(f"{'='*50}")

        results = process_response(row)
        existing_row = find_analysis_row_num(analysis_ws, idx)
        write_analysis_row(analysis_ws, idx, results, update_row=existing_row)

        state[state_key] = {"index": idx, "answers": current_answers}
        save_sync_state(state)

        action = "עודכן" if existing_row else "נוסף"
        print(f"✅ [מקומי {idx}] — {action} ב-Google Sheets")

    rows = read_responses(responses_ws)
    if not rows:
        print("לא נמצאו תגובות ב-Sheet 1.")
        return

    print(f"נמצאו {len(rows)} תגובות ייחודיות.\n")

    for i, row in enumerate(rows, start=1):
        email_col = _find_col(row, EMAIL_COL_CANDIDATES)
        email     = str(row.get(email_col, "")).strip() if email_col else f"_row_{i}"
        name_col  = _find_col(row, NAME_COL_CANDIDATES)
        name      = str(row.get(name_col, email)).strip() if name_col else email

        # --force-index N: skip all other participants
        if force_index is not None and i != force_index:
            continue

        # Check if unchanged since last sync (skip unless forced)
        current_answers = {q: str(row.get(q, "")).strip() for q in QUESTIONS}
        prev = state.get(email, {})
        if force_index is None and prev.get("answers") == current_answers:
            print(f"⏭️  [{i}] {name} — לא השתנה, מדלג")
            continue

        print(f"\n{'='*50}")
        print(f"מעבד [{i}] {name}{' (כפוי)' if force_index else ''}")
        print(f"{'='*50}")

        results = process_response(row)

        existing_row = find_analysis_row_num(analysis_ws, i)
        write_analysis_row(analysis_ws, i, results, update_row=existing_row)

        state[email] = {"index": i, "answers": current_answers}
        save_sync_state(state)

        action = "עודכן" if existing_row else "נוסף"
        print(f"✅ [{i}] {name} — {action} ב-Google Sheets")

    print("\n✅ סנכרון הושלם.")


# ─── Aggregate command ───────────────────────────────────────────────────────
MERGER_SYSTEM = """אתה מסכם משוב קבוצתי על מערכת לימוד מבוססת AI בשם Tov-Learn.
קיבלת טענות מנותחות ממספר משתתפים לאותה שאלה. תפקידך לאחד טענות דומות ולספור כמה משתתפים ציינו כל טענה.

כללים:
- מזג טענות שמבטאות את אותו רעיון, גם אם הניסוח שונה — כתוב ניסוח ברור ומייצג.
- הוסף בסוגריים כמה משתתפים ציינו כל טענה אם יותר מאחד — למשל: (ציינו 3 משתתפים).
- טענה שרק משתתף אחד ציין — אל תוסיף לה ספירה, אבל אל תשמיט אותה: כל טענה ייחודית חייבת להופיע בסיכום. אל תציין לגביה "ציין משתתף X" או מספר/זהות של משתתף ספציפי — טענה ייחודית מופיעה בלי שום תיוג, בדיוק כמו טענה עם ספירה מופיעה רק עם המספר הכולל (ציינו N משתתפים), לא עם מספור או זיהוי של משתתפים ספציפיים.
- מזג רק טענות שמבטאות בדיוק את אותו רעיון לגבי אותה ישות/סיבה. אל תמזג טענות שעוסקות בישויות שונות (למשל: תכונה אישית של משתתף לעומת מגבלה של ה-AI או של המערכת), גם אם הן מוזכרות באותו הקשר או נראות קשורות.
- אם מיזוג של טענה ייחודית לתוך טענה כללית יותר גורם לאובדן פרט (כמו תכונה אישית, סיבה ספציפית, או דוגמה) שהיה בטענה המקורית — אל תמזג. השאר את הטענה הכללית (עם ספירה אם רלוונטי) ואת הטענה הייחודית כשתי נקודות נפרדות, בלי תיוג משתתפים.
- סדר את כל הטענות (גם עם ספירה וגם בלי) ברשימה אחת, לפי פופולריות יורדת — קודם הטענות שצוינו על ידי הכי הרבה משתתפים, ואז הטענות הייחודיות.
- שמור על שפה יומיומית ופשוטה בעברית.
- כתוב בעברית בלבד — אסור להשתמש בתווים ערביים.
- כאשר הטקסט המקורי כולל מונח בשפה אחרת — שמור על הכתיב המקורי.
- אל תמציא מידע שלא נכתב."""

CATEGORY_MERGER_SYSTEM = """אתה מסכם משוב קבוצתי על מערכת לימוד מבוססת AI בשם Tov-Learn.
קיבלת טענות ממשתתפים שונים לאותה קטגוריה — שיכולה לכלול יותר משאלה אחת.
תפקידך לאחד טענות דומות ולספור כמה משתתפים ייחודיים ציינו כל טענה.

כללים:
- אם אותו משתתף ציין את אותה טענה בשתי שאלות שונות — ספר אותה פעם אחת עבור אותו משתתף.
- מזג טענות שמבטאות את אותו רעיון ממשתתפים שונים — כתוב ניסוח ברור ומייצג.
- הוסף בסוגריים כמה משתתפים ייחודיים ציינו כל טענה, אם יותר מאחד — למשל: (ציינו 3 משתתפים).
- טענה שרק משתתף אחד ציין — אל תוסיף לה ספירה, אבל אל תשמיט אותה: כל טענה ייחודית חייבת להופיע בסיכום. אל תציין לגביה "ציין משתתף X" או מספר/זהות של משתתף ספציפי — טענה ייחודית מופיעה בלי שום תיוג, בדיוק כמו טענה עם ספירה מופיעה רק עם המספר הכולל (ציינו N משתתפים), לא עם מספור או זיהוי של משתתפים ספציפיים.
- מזג רק טענות שמבטאות בדיוק את אותו רעיון לגבי אותה ישות/סיבה. אל תמזג טענות שעוסקות בישויות שונות (למשל: תכונה אישית של משתתף לעומת מגבלה של ה-AI או של המערכת), גם אם הן מוזכרות באותו הקשר או נראות קשורות.
- אם מיזוג של טענה ייחודית לתוך טענה כללית יותר גורם לאובדן פרט (כמו תכונה אישית, סיבה ספציפית, או דוגמה) שהיה בטענה המקורית — אל תמזג. השאר את הטענה הכללית (עם ספירה אם רלוונטי) ואת הטענה הייחודית כשתי נקודות נפרדות, בלי תיוג משתתפים.
- סדר את כל הטענות (גם עם ספירה וגם בלי) ברשימה אחת, לפי פופולריות יורדת — קודם הטענות שצוינו על ידי הכי הרבה משתתפים, ואז הטענות הייחודיות.
- שמור על שפה יומיומית ופשוטה בעברית.
- כתוב בעברית בלבד — אסור להשתמש בתווים ערביים.
- כאשר הטקסט המקורי כולל מונח בשפה אחרת — שמור על הכתיב המקורי.
- אל תמציא מידע שלא נכתב."""

CATEGORY_GROUPS = [
    {
        "name": "חיובי — מה עבד טוב",
        "questions": ["מה עבד טוב?", "אילו חלקים היו שימושיים במיוחד?"],
    },
    {
        "name": "שלילי — קשיים ושיפורים",
        "questions": ["מה היה לא ברור או קשה?", "מה חשוב לתקן/לשפר? מה היה חסר?"],
    },
    {
        "name": "רעיונות חדשים",
        "questions": ["האם יש לכם רעיונות חדשים להוסיף?"],
    },
]

_SKIP_CELL = {"", "לא ענה"}   # cell values to always exclude from aggregate


def _strip_preserved(text: str) -> str:
    """Remove (*old value*) sections kept from previous analyses."""
    return re.sub(r'\n\n\(\*.*?\*\)', '', text, flags=re.DOTALL).strip()


def _cell_is_aggregate_worthy(raw: str) -> bool:
    """True only when the cell has actual analyzed claims."""
    s = raw.strip()
    return bool(s) and s not in _SKIP_CELL and not s.startswith("⚠️") and not s.startswith("—")


def aggregate_command():
    print("מתחבר ל-Google Sheets...")
    spreadsheet, _, analysis_ws = get_sheets()

    raw_records = analysis_ws.get_all_records()
    # Keep only rows where "אינדקס" is a valid integer (filters stray manual rows)
    all_records = [r for r in raw_records
                   if str(r.get("אינדקס", "")).strip().lstrip("-").isdigit()]

    if not all_records:
        print("Sheet 2 ריק — הרץ sync קודם.")
        return

    skipped = len(raw_records) - len(all_records)
    if skipped:
        print(f"⚠️  דולג על {skipped} שורות עם אינדקס לא תקין — בדוק ומחק אותן ב-Sheet 2.\n")

    print(f"מאחד טענות מ-{len(all_records)} משתתפים...\n")

    # ── Part 1: per-question summary ──────────────────────────────────────────
    print("חלק 1: סיכום לפי שאלה")
    question_rows = [["שאלה", "סיכום משולב"]]

    for q in QUESTIONS:
        cells = [_strip_preserved(str(r.get(q, "")))
                 for r in all_records
                 if _cell_is_aggregate_worthy(str(r.get(q, "")))]

        print(f"  {q[:38]}... ({len(cells)} תגובות)")

        if not cells:
            question_rows.append([q, "אין נתונים."])
            continue
        if len(cells) == 1:
            question_rows.append([q, cells[0]])
            continue

        labeled = "\n\n".join(f"[משתתף {i+1}]\n{c}" for i, c in enumerate(cells))
        prompt = (
            f"שאלה: {q}\n\n"
            f"להלן תשובות מנותחות של {len(cells)} משתתפים (כל נקודת בלט = טענה אחת):\n\n"
            f"{labeled}\n\n"
            f"אחד את כל הטענות: מצא טענות זהות או דומות, מזג אותן, "
            f"הוסף כמה משתתפים ציינו כל אחת (אם יותר מאחד) וסדר לפי פופולריות."
        )
        question_rows.append([q, call_api(MERGER_SYSTEM, prompt)])
        print(f"    ✅ אוחדו")

    # ── Part 2: category summary ───────────────────────────────────────────────
    print("\nחלק 2: סיכום לפי קטגוריה")
    category_rows = [["קטגוריה", "סיכום משולב"]]

    for group in CATEGORY_GROUPS:
        gname = group["name"]
        items = [
            (p_num, q, _strip_preserved(str(record.get(q, ""))))
            for p_num, record in enumerate(all_records, start=1)
            for q in group["questions"]
            if _cell_is_aggregate_worthy(str(record.get(q, "")))
        ]

        unique_participants = len(set(p for p, _, _ in items))
        print(f"  {gname[:38]}... ({unique_participants} משתתפים, {len(items)} תאים)")

        if not items:
            category_rows.append([gname, "אין נתונים."])
            continue
        if unique_participants == 1:
            category_rows.append([gname, "\n".join(c for _, _, c in items)])
            continue

        by_p = defaultdict(list)
        for p_num, q, cell in items:
            by_p[p_num].append((q, cell))

        labeled_parts = [
            f"[משתתף {p} | {q}]\n{cell}"
            for p in sorted(by_p)
            for q, cell in by_p[p]
        ]
        prompt = (
            f"קטגוריה: {gname}\n\n"
            f"להלן טענות ממשתתפים שונים:\n\n"
            f"{chr(10).join(labeled_parts)}\n\n"
            f"הנחיות:\n"
            f"1. אם אותו משתתף ציין את אותה טענה בשתי שאלות — ספר אותה פעם אחת.\n"
            f"2. מזג טענות דומות ממשתתפים שונים לטענה אחת.\n"
            f"3. הוסף כמה משתתפים ייחודיים ציינו כל טענה (אם יותר מאחד).\n"
            f"4. סדר לפי פופולריות יורדת."
        )
        category_rows.append([gname, call_api(CATEGORY_MERGER_SYSTEM, prompt)])
        print(f"    ✅ אוחדו")

    # ── Write to "סיכום" tab ───────────────────────────────────────────────────
    all_rows = (
        [["", "סיכום לפי שאלה"]]
        + question_rows
        + [["", ""], ["", "סיכום לפי קטגוריה"]]
        + category_rows
    )

    try:
        summary_ws = spreadsheet.worksheet("סיכום")
        summary_ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        summary_ws = spreadsheet.add_worksheet(title="סיכום", rows=30, cols=2)

    summary_ws.update(
        range_name="A1",
        values=all_rows,
        value_input_option="USER_ENTERED"
    )
    print("\n✅ הסיכום נכתב ל-Tab 'סיכום' ב-Google Sheets.")


# ─── Entry point ──────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="כלי לניתוח משובים על Tov-Learn")
    parser.add_argument(
        "mode",
        choices=["validate", "analyze", "sync", "aggregate"],
        help="validate / analyze — עובד על feedbacks.csv | sync — Google Sheets | aggregate — סיכום משולב"
    )
    parser.add_argument("--index",        type=int,  default=0,     help="אינדקס המשתתף (ברירת מחדל: 0)")
    parser.add_argument("--all",          action="store_true",       help="עבד את כל השורות")
    parser.add_argument("--list-columns", action="store_true",       help="(sync בלבד) הצג שמות עמודות ויצא")
    parser.add_argument("--force-index",  type=int, default=None,    help="(sync בלבד) עבד מחדש משתתף לפי אינדקס, גם אם לא השתנה")
    args = parser.parse_args()

    if args.mode == "sync":
        sync_command(list_columns=args.list_columns, force_index=args.force_index)
        return

    if args.mode == "aggregate":
        aggregate_command()
        return

    df = pd.read_csv("feedbacks.csv", index_col=0, encoding="utf-8-sig")

    if args.all:
        for idx in df.index:
            if args.mode == "validate":
                validate_row(idx, df)
            else:
                analyze_row(idx, df)
    else:
        if args.mode == "validate":
            validate_row(args.index, df)
        else:
            analyze_row(args.index, df)


if __name__ == "__main__":
    main()
