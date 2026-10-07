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

תפקידך: לבדוק אם תשובה במשוב ניתנת לניתוח — כלומר, האם אפשר להבין ממנה מה המשתתף טוען ועל מה.
דוגמאות, סיפורים או פירוט רב הם יתרון, אבל אינם תנאי. תשובה קצרה וכללית תקפה לגמרי אם היא אומרת משהו אמיתי על נושא מזוהה.

תשובה מספקת אם מתקיימים בה כל התנאים הבאים:
1. נושא מזוהה — ברור על איזה חלק או היבט של הקורס/המערכת מדובר. נושא כללי מספיק: "הבחנים", "התרגילים", "קצב הלימוד", "השיחה עם ה-AI", "ההתקנה". אין צורך בשיעור או תרגיל מסוים.
2. טענה ממשית — נאמר משהו על הנושא: שהוא עזר, הפריע, היה קשה, היה חסר, כדאי להוסיף אותו וכו'.
   השאלה עצמה יכולה לספק את הטענה: תשובה שרק מציינת נושא בתגובה ל"מה עבד טוב?" אומרת שהנושא עבד טוב; בתגובה ל"מה היה קשה?" — שהוא היה קשה, וכן הלאה. זו טענה ממשית ומספקת.
3. מובנות — אפשר להבין את התשובה ולסכם אותה כנקודת משוב, והיא עונה על השאלה שנשאלה.
   תשובה שכולה עונה בבירור על שאלה אחרת אינה עונה על השאלה — למשל תשובה שכולה מתארת מה היה שימושי או מה עבד טוב, בתגובה ל"מה היה לא ברור או קשה?" (קורה כשמשתתף מעתיק תשובה לשדה הלא נכון).
   תשובה שעונה על השאלה וכוללת בנוסף הערה מסוג אחר — תקפה (למשל תשובה ל"מה חשוב לתקן?" שמונה כמה תיקונים ומוסיפה "היחידות מסודרות בצורה לוגית").

גם תשובת "אין" מספקת — המשתתף מציין במפורש שאין לו קשיים, רעיונות, תלונות וכו' (למשל: "לא היו לי קשיים", "הכל היה בסדר", "אין לי רעיונות להוסיף"). היא עונה על השאלה.
סמן nothing_to_report=true רק כשכל התשובה היא תשובת "אין" כזו. אם יש בתשובה ולו טענה אחת ממשית — גם אם היא מנוסחת בשלילה ("ההסברים לא היו עמוקים מספיק", "ה-JSON לא היה ברור", "לא היו מספיק תרגולים") — זו טענה, ו-nothing_to_report=false.

דוגמאות לתשובות מספקות:
- "הכל היה בסדר." / "לא נתקלתי בקשיים." / "אין לי רעיונות להוסיף." (תשובת "אין")
- "הבחנים עזרו לי לדעת כמה אני שולטת בחומר." (נושא: הבחנים; טענה: עזרו להעריך שליטה בחומר)
- "לא הבנתי את התרגילים." (נושא: התרגילים; טענה: לא היו מובנים — כללי, אבל ניתן לניתוח)
- "הקצב היה מהיר מדי." (נושא: קצב הלימוד; טענה: מהיר מדי)
- "השיחה עם Claude." בתגובה ל"מה עבד טוב?" (נושא: השיחה עם ה-AI; הטענה מגיעה מהשאלה — היא עבדה טוב)
- "הבחן בסוף שיעור 2.4 עזר לי לגלות שלא הבנתי את ensure_ascii=False — זה גרם לי לחזור ולתרגל שוב."
- "קצב וצורת לימוד מותאמים לרמה. אפשרות לדון עם המודל כדי להבין את החומר. בדיקה שהחומר הובן. התנסות בתרגילים ושיעורי בית."

תשובה אינה מספקת רק כשאי אפשר לנתח אותה, כלומר:
- אין נושא מזוהה — לא ברור על מה מדובר: "חלק מהדברים היו קשים.", "היו כמה בעיות.", "אחד השיעורים היה מבלבל."
- אין טענה ממשית — לא ניתן לדעת מה המשתתף אומר, גם בהתחשב בשאלה: "בערך.", "תלוי.", "לא בטוחה."
- לא מובנת — משובשת, סותרת את עצמה, או מפנה למשהו שאינו זמין: "כמו שכתבתי קודם.", "זה שאמרתי לך."
- עונה על שאלה אחרת — כל התשובה שייכת בבירור לשאלה אחרת: "הנושאים שהיו שימושיים במיוחד עבורי הם Prompt Engineering ו-RAG." בתגובה ל"מה היה לא ברור או קשה?".

אם התשובה אינה מספקת, כתוב שאלת המשך קצרה וספציפית בעברית שמכוונת בדיוק למה שחסר (הנושא, הטענה, או הבהרה).
לתשובה שעונה על שאלה אחרת — ציין בעדינות שנראה שהתשובה מתאימה לשאלה אחרת, ושאל שוב את השאלה המקורית.
בשדה reason כתוב בקצרה בעברית מה הנושא והטענה שזיהית, או מה חסר.
החזר JSON בלבד, ללא טקסט נוסף: {"sufficient": true/false, "nothing_to_report": true/false, "reason": "הסבר קצר", "followup": "שאלת המשך" או null}"""

ANALYZER_SYSTEM = """אתה מנתח משוב על מערכת לימוד מבוססת AI בשם Tov-Learn.
קיבלת תשובה של אחד המשתתפים לשאלה אחת. תפקידך לסכם את הנקודות המרכזיות לשאלה זו.

כללים בסיסיים:
- הישאר צמוד למה שנכתב — אל תמציא מידע ואל תשנה את משמעות הטענות.
- סכם בעברית יומיומית ופשוטה — אל תשתמש במילים פורמליות או אקדמיות.
- כתוב בעברית בלבד — אסור להשתמש בתווים ערביים.
- כאשר הטקסט המקורי כולל מונח, ביטוי, או שם בשפה אחרת (כמו "compact", "Vibe Coding", "Lesson Summaries", "Lesson") — שמור על הכתיב המקורי. אל תתעתק אותו לעברית. לדוגמה: "compact" נשאר "compact" ולא "קומפקט"; "Vibe Coding" נשאר "Vibe Coding" ולא "ויב קודינג".
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

קושי אישי מול גורם אנושי, ולמה ה-AI לא יכול היה לעזור בו:
אם המשתתף מתאר קושי אישי שלו (כמו קשיי תקשורת, נטייה לפרש לא נכון) שגרם לו קושי להבין דבר-מה לבד או מול גורם אנושי (כמו מנחים), ומסביר שה-AI לא יכול היה לעזור בזה כי אין לו גישה לכוונות של אותו גורם — ההסבר על ה-AI הוא הנימוק של אותו קושי, לא טענה נפרדת ולא תקלה של ה-AI שצריך לתקן. כתוב נקודה אחת שכוללת את הקושי ואת ההסבר.
בתוך הנקודה שמור על ייחוס נכון: התכונה האישית שייכת למשתתף, וה-AI לא יכול היה לעזור מסיבה מבנית (אין לו את כוונות הגורם החיצוני) — לא בגלל התכונה של המשתתף. השוואה לעבודה אמיתית שמסבירה את זה — חלק מאותה נקודה.
גם הערה שמבהירה את גבולות הבעיה (למשל "שאלות על החומר עצמו יכולתי לשאול את ה-AI, הבעיה הייתה רק עם מה שמחוץ להקשר שלו") — חלק מאותה נקודה, לא טענה חיובית נפרדת.
שגוי — מייחס את קשיי התקשורת ל-AI:
- היה קשה לשאול את ה-AI שאלות שלא קשורות לחומר, כי ה-AI לא ידע מה המנחים התכוונו, בשל קשיי תקשורת ונטייה לפרש דברים לא נכון.
שגוי — מפצל את הנימוק לנקודה נפרדת, כאילו זו תקלה של ה-AI:
- היה קשה להבין לבד הגדרות שקבעו המנחים, כי יש לה קשיי תקשורת ונטייה לפרש דברים לא נכון.
- ה-AI לא ידע לענות על שאלות שמחוץ להקשר של החומר, כי הוא לא יכול לדעת מה המנחים התכוונו.
נכון — נקודה אחת:
- היה לה קשה עם הגדרות שהמנחים קבעו לפרויקטים שאינן חלק מהחומר: המנחים רצו שתבין אותן באופן עצמאי, אבל לא תמיד הבינה לבד, כי יש לה קשיי תקשורת ונטייה לפרש דברים לא נכון. ה-AI לא יכול היה לעזור בזה, כי אין לו גישה למה שהמנחים התכוונו — שאלות על החומר עצמו כן יכלה לשאול אותו — בדיוק כמו שבפרויקטים אמיתיים רק המנהל או הלקוח שביקשו את הפרויקט יודעים מה עבר בראש שלהם.

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
נכון: "...שהיה לוקח הרבה זמן"

דברים שכבר תוקנו — השמט:
טענה שהמשתתף מציין במפורש שכבר תוקנה או נפתרה ("תוקן", "זה כבר סודר", "בהתחלה X, אבל זה השתנה") — השמט אותה לגמרי. היא כבר לא משוב שצריך לפעול לפיו.
אם המשתתף פותח בהקשר כמו "דברים שכבר תוקנו במהלך הלימודים" — כל מה שנכתב באותו הקשר ומתאר מצב שכבר תוקן, גם בלי המילה "תוקן", מושמט. טענה באותה תשובה שמתארת בעיה שעדיין קיימת (בזמן הווה, למשל "אין הוראות ל...") — נשארת.

טענה מקטגוריה אחרת:
לכל שאלה יש קטגוריה — חיובי (מה עבד טוב / מה היה שימושי), שלילי (קשיים ומה לתקן) או רעיון חדש. ניתנת לך הקטגוריה של השאלה.
- טענה שמתאימה לקטגוריה של השאלה — ב-claims.
- טענה שבבירור שייכת לקטגוריה אחרת (למשל טענה חיובית בתשובה לשאלה על קשיים, או בקשה לתקן משהו קיים בתשובה לשאלת הרעיונות) ועדיין רלוונטית — ב-other_category, עם הקטגוריה המתאימה: "positive", "negative" או "idea".
- טענה מקטגוריה אחרת שלפי ההקשר כבר לא רלוונטית — למשל טענה חיובית שמתארת את התוצאה של משהו שכבר תוקן — השמט אותה.
- נימוק או הבהרה שהם חלק מטענה (ראה למעלה) — אינם טענה מקטגוריה אחרת; הם נשארים בתוך הטענה.
דוגמה — תשובה ל"מה חשוב לתקן/לשפר?" (שלילי): "דברים שכבר תוקנו במהלך הלימודים, אז אפרט בקצרה. בהתחלה המורה לא הציג תרגולים - תוקן. חלק מהחומר לא היה עדכני - תוקן. היחידות מסודרות בצורה לוגית על פי רוב. אין הוראות למורה לוודא ולעדכן את החומר לפני הצגתו."
נכון: claims = ["אין הוראות למורה לוודא ולעדכן את החומר לפני הצגתו."], other_category = []
(שתי הבעיות הראשונות כבר תוקנו → מושמטות. "היחידות מסודרות בצורה לוגית" חיובית, ובהקשר של "דברים שכבר תוקנו" מתארת מצב שכבר תוקן → מושמטת. הבעיה האחרונה עדיין קיימת → נשארת.)

אם לא נשארה אף טענה (הכול כבר תוקן או לא רלוונטי) — החזר רשימות ריקות.

פורמט הפלט — JSON בלבד:
{"claims": ["טענה אחת", "טענה אחרת"], "other_category": [{"category": "positive", "text": "טענה"}]}
כל טענה היא מחרוזת אחת, בלי תו בלט בתחילתה."""


# ─── Core API helpers ─────────────────────────────────────────────────────────
# Letters from scripts other than Hebrew/Latin (e.g. Arabic "ال", Georgian "ექ") that the
# model sometimes slips into Hebrew text. Digits, punctuation and symbols are not matched.
_FOREIGN_LETTERS = re.compile(r"[^\W\d_a-zA-Z֐-׿]")


def call_api(system_prompt: str, user_prompt: str, json_mode: bool = False,
             temperature: float | None = None) -> str:
    if temperature is None:
        temperature = 0.0 if json_mode else 0.3   # JSON = merge/classify: favor stability
    from google.genai.errors import ClientError
    for attempt in range(4):
        try:
            response = client.models.generate_content(
                model=MODEL,
                contents=user_prompt,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    temperature=temperature,
                    response_mime_type="application/json" if json_mode else None,
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


def parse_json(text: str) -> dict:
    match = re.search(r'\{.*?\}', text, re.DOTALL)
    if not match:
        return {"sufficient": True, "followup": None}
    try:
        return json.loads(match.group())
    except json.JSONDecodeError:
        return {"sufficient": True, "followup": None}


def validate_answer(q: str, answer: str) -> dict:
    """Ask the validator whether one answer can be analyzed.
    Returns {sufficient, nothing_to_report, reason, followup}."""
    prompt = (
        f'שאלה: "{q}"\n'
        f'תשובה: "{answer}"\n\n'
        f'האם התשובה ניתנת לניתוח — יש בה נושא מזוהה וטענה ממשית, והיא מובנת? והאם כל התשובה היא תשובת "אין"?\n'
        f'החזר JSON בלבד: {{"sufficient": true/false, "nothing_to_report": true/false, "reason": "הסבר קצר", "followup": "שאלת המשך או null"}}'
    )
    return parse_json(call_api(VALIDATOR_SYSTEM, prompt))


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

        result = validate_answer(q, answer)
        reason = result.get("reason") or ""

        if result.get("sufficient"):
            print(f"✅ {q}\n    {reason or 'התשובה ניתנת לניתוח.'}\n")
        else:
            followup = result.get("followup", "")
            print(f"❌ {q}\n    {reason}\n    שאלת המשך לשליחה: {followup}\n")
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

    for q in QUESTIONS:
        answer = str(row.get(q, "")).strip()
        print(f"--- {q}")
        if not answer or answer == "nan":
            print("לא ענה\n")
            continue
        print(f"{analyze_answer(q, answer) or '(כל הנקודות כבר תוקנו או אינן רלוונטיות עוד)'}\n")


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
        val = validate_answer(q, answer)

        if not val.get("sufficient"):
            followup = val.get("followup", "")
            results[q] = f"⚠️ תשובה לא מספיקה\nשאלת המשך: {followup}\n\n{answer}"
            print(f"  ❌ {q[:40]}...")
        elif val.get("nothing_to_report"):
            # Valid "nothing to report" (judged by the validator, not by keywords —
            # negatively-phrased claims like "X לא היה ברור" are real claims) —
            # readable in Sheet 2, filtered from aggregate
            results[q] = "— " + answer
            print(f"  — {q[:40]}... (אין נתונים לדיווח)")
        else:
            # Step 2 — analyze this question alone
            cell = analyze_answer(q, answer)
            if cell is None:
                # Every claim was already fixed or no longer relevant — nothing to act on.
                # The raw answer stays readable; '—' keeps it out of the aggregate.
                results[q] = f"— {answer}\n(כל הנקודות כבר תוקנו או אינן רלוונטיות עוד)"
                print(f"  — {q[:40]}... (הכול כבר תוקן / לא רלוונטי)")
            else:
                results[q] = cell
                print(f"  ✅ {q[:40]}...")

    return results


# Off-category claims (e.g. a positive remark in an answer to a negative question) are
# kept in the analysis cell under this header, tagged with their category label. The
# per-question summary skips them; the category summary files them by their tag.
OTHER_CATEGORY_HEADER = "שייך לקטגוריה אחרת (מופיע רק בסיכום לפי קטגוריה):"
CATEGORY_LABELS = {"positive": "חיובי", "negative": "שלילי", "idea": "רעיון חדש"}


def _question_category(q: str) -> str:
    return next(g["key"] for g in CATEGORY_GROUPS if q in g["questions"])


def analyze_answer(q: str, answer: str) -> str | None:
    """Analyze one valid answer into an analysis-tab cell: one bullet per claim, plus an
    OTHER_CATEGORY_HEADER section for still-relevant claims of another category.
    Claims that were already fixed, or are no longer relevant, are dropped by the analyzer.
    Returns None if no claim is left."""
    key = _question_category(q)
    prompt = (f"שאלה: {q}\n"
              f"קטגוריית השאלה: {key} ({CATEGORY_LABELS[key]})\n"
              f"תשובה: {answer}")

    def check(result) -> str | None:
        if not isinstance(result, dict) or not isinstance(result.get("claims", []), list):
            return "מבנה תשובת הניתוח שגוי"
        bad = [o for o in result.get("other_category", [])
               if not isinstance(o, dict) or o.get("category") not in CATEGORY_LABELS]
        return "קטגוריה לא מוכרת בניתוח" if bad else None

    result = _call_json(ANALYZER_SYSTEM, prompt, key=None, check=check)
    if not isinstance(result, dict):
        print("  ⚠️ הניתוח נכשל — נשמרת התשובה המקורית כנקודה אחת")
        return f"- {answer}"

    def clean(text) -> str:   # strip a leading bullet; foreign letters only as a last resort
        return _FOREIGN_LETTERS.sub("", re.sub(r"^[-*•]\s*", "", str(text))).strip()

    claims = [clean(c) for c in result.get("claims", []) if clean(c)]
    other = []
    for o in result.get("other_category", []):
        if not isinstance(o, dict) or not clean(o.get("text", "")):
            continue
        if o.get("category") == key:      # not actually another category
            claims.append(clean(o["text"]))
        elif o.get("category") in CATEGORY_LABELS:
            other.append((o["category"], clean(o["text"])))

    lines = [f"- {c}" for c in claims]
    if other:
        lines += ([""] if lines else []) + [OTHER_CATEGORY_HEADER]
        lines += [f"- [{CATEGORY_LABELS[cat]}] {text}" for cat, text in other]
    return "\n".join(lines) or None


def find_analysis_row_num(analysis_ws, index: int) -> int | None:
    """Return 1-based row number in Sheet 2 where column A == index, or None."""
    col_a = analysis_ws.col_values(1)
    for i, val in enumerate(col_a):
        if str(val) == str(index):
            return i + 1
    return None


def _split_preserved(text: str) -> tuple[str, str | None]:
    """Split a cell into (current value, preserved last-valid version or None)."""
    current, sep, rest = text.partition("\n\n(*")
    if not sep:
        return text.strip(), None
    preserved = rest[:-2] if rest.endswith("*)") else rest
    return current.strip(), preserved.strip() or None


def _is_valid_value(value: str) -> bool:
    """A cell value holding a valid answer: analyzed claims or a '—' nothing-to-report."""
    s = value.strip()
    return bool(s) and not s.startswith("⚠️") and s.lstrip("-* ").strip() not in _SKIP_CELL


def write_analysis_row(analysis_ws, index: int, results: dict[str, str], update_row: int | None,
                       changed: set[str] | None = None):
    """Append a new row, or update an existing one.
    History is kept only when the participant replaced a valid answer with a new invalid
    one: the new ⚠️ value keeps the last valid version in (*...*). `changed` = questions
    whose answer text changed since the last sync (None = treat all as changed).
    If the answer is unchanged and only re-judged invalid (e.g. after a stricter
    validation rule), the old analysis is not kept — it analyzed this same answer.
    An already-kept version (from an earlier, different answer) is carried forward.
    Any other update replaces the cell outright."""
    row_values = [str(index)] + [results.get(q, "") for q in QUESTIONS]
    columns = [None] + QUESTIONS

    if update_row is None:
        analysis_ws.append_row(row_values, value_input_option="USER_ENTERED")
        return

    # Update — merge with existing values
    existing = analysis_ws.row_values(update_row)
    while len(existing) < len(row_values):
        existing.append("")

    merged = []
    for q, new_val, old_val in zip(columns, row_values, existing):
        if not new_val.startswith("⚠️"):
            merged.append(new_val)
            continue
        old_current, old_preserved = _split_preserved(old_val)
        answer_changed = changed is None or q in changed
        if answer_changed and _is_valid_value(old_current):
            last_valid = old_current
        else:
            last_valid = old_preserved
        merged.append(f"{new_val}\n\n(*{last_valid}*)" if last_valid else new_val)

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
        prev_answers = state.get(state_key, {}).get("answers") or {}
        changed = {q for q in QUESTIONS if prev_answers.get(q) != current_answers[q]}
        write_analysis_row(analysis_ws, idx, results, update_row=existing_row, changed=changed)

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
        prev_answers = prev.get("answers") or {}
        changed = {q for q in QUESTIONS if prev_answers.get(q) != current_answers[q]}
        write_analysis_row(analysis_ws, i, results, update_row=existing_row, changed=changed)

        state[email] = {"index": i, "answers": current_answers}
        save_sync_state(state)

        action = "עודכן" if existing_row else "נוסף"
        print(f"✅ [{i}] {name} — {action} ב-Google Sheets")

    print("\n✅ סנכרון הושלם.")


# ─── Aggregate command ───────────────────────────────────────────────────────
_MERGE_RULES = """כל טענה בקלט מסומנת במזהה בסוגריים מרובעים, למשל [p3.2] = משתתף 3, טענה 2.
תפקידך לאחד טענות זהות או דומות לטענה ממוזגת אחת, ולציין ב-sources את המזהים של כל הטענות שמוזגו לתוכה.
אינך סופר משתתפים ואינך כותב ספירות — הקוד מחשב כמה משתתפים ציינו כל טענה מתוך sources.

כללים:
- כל מזהה מהקלט חייב להופיע ב-sources של בדיוק טענה ממוזגת אחת. אל תשמיט אף טענה, גם אם רק משתתף אחד ציין אותה.
- מזג רק טענות שמבטאות בדיוק את אותו רעיון לגבי אותה ישות, בעיה או סיבה. שתי בעיות שונות לא מתמזגות, גם אם הן מופיעות באותו הקשר, נוגעות לאותו רכיב, או נשמעות קשורות.
  דוגמה: "חשוב להוסיף אפשרות לעבור שקופית שקופית, כי המערכת הציגה כמה שקופיות בבת אחת" ו"חשוב לדאוג שהמערכת לא תדלג על תרגולים" — שתי בעיות שונות → שתי טענות נפרדות.
- אל תמזג טענות שעוסקות בישויות שונות (למשל: תכונה אישית של משתתף לעומת מגבלה של ה-AI או של המערכת), גם אם הן מוזכרות באותו הקשר.
  דוגמה: "היה לה קשה להבין לבד הגדרות של המנחים, כי יש לה קשיי תקשורת" ו"ההסברים לא היו עמוקים מספיק" — קושי שנובע מתכונה אישית מול מגבלה של המערכת → שתי טענות נפרדות.
- תכונה אישית של משתתף נשארת שלו בלבד — אל תכליל אותה ל"המשתתפים" ואל תייחס אותה ל-AI או למערכת. שמור על ניסוח הבעלות כמו במקור: "ויש לה קשיי תקשורת", לא "ויש קשיי תקשורת" או "היו קשיי תקשורת".
- שתי טענות שמבטאות את אותו רעיון מרכזי, כשאחת מפורטת יותר מהשנייה — מזג אותן, וכתוב את הטענה הממוזגת כך שתכלול את כל הפרטים של שתיהן (תכונה אישית, סיבה ספציפית, דוגמה). אל תשמיט פרט, ואל תשאיר אותן נפרדות רק בגלל שרמת הפירוט שונה.
  דוגמה: "האפשרות לשאול שאלות על מה שלא הבינה, כך שה-AI ענה מתוך ההקשר של החומר, עזרה להבין באופן עצמאי, במקום לשאול את המנחים" ו"האפשרות לשאול שאלות בפירוט והתשובות שהמערכת מציעה" — אותו רעיון (האפשרות לשאול את ה-AI שאלות) → טענה אחת שכוללת את הפרטים של שתיהן.
- לעולם אל תשים ב-sources של טענה אחת שני מזהים של אותו משתתף שנכתבו תחת אותה שאלה — כל מזהה כזה הוא כבר טענה נפרדת.
- טענות של אותו משתתף מתמזגות רק אם הן אותה טענה בדיוק שנכתבה תחת שתי שאלות שונות. טענות שונות של אותו משתתף נשארות נפרדות, גם אם הן באותו נושא כללי.
  דוגמה: "בהתחלה המורה לא הציג תרגולים (תוקן)" ו"חלק מהחומר לא היה עדכני (תוקן)" — שתי בעיות שונות → שתי טענות נפרדות.
- כשאתה בספק אם שתי טענות הן אותו רעיון — אל תמזג. רמת פירוט שונה של אותו רעיון אינה סיבה לספק.
- הטקסט של כל טענה ממוזגת הוא נקודה אחת, בלי ספירה, בלי מספר משתתפים, בלי מזהים, בלי תיוג משתתף ("ציין משתתף X"), ובלי הקדמה או סיכום.
- שמור על שפה יומיומית ופשוטה בעברית.
- כתוב בעברית בלבד — אסור להשתמש בתווים ערביים.
- כאשר הטקסט המקורי כולל מונח בשפה אחרת — שמור על הכתיב המקורי.
- אל תמציא מידע שלא נכתב."""

MERGER_SYSTEM = f"""אתה מסכם משוב קבוצתי על מערכת לימוד מבוססת AI בשם Tov-Learn.
קיבלת טענות מנותחות ממספר משתתפים לאותה שאלה או לאותה קטגוריה.

{_MERGE_RULES}

החזר JSON בלבד: {{"claims": [{{"text": "ניסוח הטענה הממוזגת", "sources": ["p3.2", "p5.1"]}}]}}"""

CATEGORY_MERGER_SYSTEM = """אתה מסכם משוב קבוצתי על מערכת לימוד מבוססת AI בשם Tov-Learn.
קיבלת טענות שכבר אוחדו בסיכום לפי שאלה, וכולן שייכות לאותה קטגוריה. כל טענה מסומנת במזהה [uN], ולידה השאלה שתחתיה נכתבה והמשתתפים שציינו אותה.
תפקידך לאחד טענות זהות שעדיין מופיעות בנפרד — בעיקר אותו רעיון שנכתב תחת שאלות שונות (למשל פעם תחת "מה עבד טוב?" ופעם תחת "אילו חלקים היו שימושיים?"), וגם אותו רעיון ממשתתפים שונים תחת אותה שאלה, אם לא אוחד קודם.

כללים:
- לעולם אל תמזג שתי טענות מאותה שאלה שיש להן משתתף משותף — אלה טענות שונות של אותו משתתף.
- מזג רק טענות שמבטאות בדיוק את אותו רעיון לגבי אותה ישות, בעיה או סיבה. כשאתה בספק — אל תמזג.
- כל מזהה מהקלט חייב להופיע ב-sources של בדיוק טענה אחת. טענה שלא מוזגה עם אף אחת — החזר אותה לבד, בניסוח המקורי.
- בטענה ממוזגת כתוב ניסוח ברור שמכיל את הפרטים של כל הטענות שמוזגו — אל תשמיט פרט.
- אל תכתוב בטקסט ספירה, מספר משתתפים, מזהים, תיוג משתתף, הקדמה או סיכום — הקוד מוסיף את הספירה.
- תכונה אישית של משתתף נשארת שלו בלבד ובניסוח הבעלות המקורי ("ויש לה קשיי תקשורת").
- כתוב בעברית בלבד — אסור להשתמש בתווים ערביים. מונחים בשפה אחרת — בכתיב המקורי. אל תמציא מידע שלא נכתב.

החזר JSON בלבד: {"claims": [{"text": "ניסוח הטענה", "sources": ["u3", "u17"]}]}"""

TOPIC_GROUPER_SYSTEM = """אתה עוזר לסדר סיכום משוב על מערכת לימוד מבוססת AI בשם Tov-Learn.
קיבלת רשימה ממוספרת של טענות מתוך הסיכום, וכולן מאותה קטגוריה. קבץ אותן לקבוצות לפי נושא דומה או קשור, כדי שאדם יוכל לעבור על כל קבוצה ולהחליט בעצמו אילו טענות למזג.

כללים:
- קבוצה = טענות שעוסקות באותו נושא או בנושאים קשורים — למשל כל הטענות על תרגול מעשי, כל הטענות על שאילת שאלות ל-AI, כל הטענות על שיעורים ונושאים שימושיים ספציפיים, כל הטענות על בעיות עם עברית ותצוגה.
- טענות קשורות נכנסות לאותה קבוצה גם אם הן לא אותה טענה — זה רק קיבוץ לעיון, לא מיזוג.
- טענה שאין לה נושא משותף עם אף טענה אחרת — קבוצה של טענה אחת.
- כל מספר מהרשימה חייב להופיע בדיוק בקבוצה אחת.

החזר JSON בלבד: {"groups": [[1, 4], [2], [3, 5, 6]]}"""

# A claim belongs to its question's category, unless the analyzer tagged it as another
# category (OTHER_CATEGORY_HEADER section of the analysis cell).
CATEGORY_GROUPS = [
    {
        "key": "positive",
        "name": "חיובי — מה עבד טוב",
        "questions": ["מה עבד טוב?", "אילו חלקים היו שימושיים במיוחד?"],
    },
    {
        "key": "negative",
        "name": "שלילי — קשיים ושיפורים",
        "questions": ["מה היה לא ברור או קשה?", "מה חשוב לתקן/לשפר? מה היה חסר?"],
    },
    {
        "key": "idea",
        "name": "רעיונות חדשים",
        "questions": ["האם יש לכם רעיונות חדשים להוסיף?"],
    },
]

_SKIP_CELL = {"", "לא ענה"}   # cell values to always exclude from aggregate


def _aggregate_value(raw: str) -> str | None:
    """The analyzed claims a cell contributes to the summary, or None.
    Uses the current value; if it's ⚠️ invalid, falls back to the kept last-valid version.
    '—' nothing-to-report values contribute nothing."""
    current, kept = _split_preserved(raw)
    value = kept if current.startswith("⚠️") else current
    if not value or not _is_valid_value(value) or value.startswith("—"):
        return None
    return value


def _split_claims(cell: str) -> list[tuple[str, str | None]]:
    """One analyzed cell -> [(claim, category key or None)].
    Main bullets get None (= the question's own category); bullets under
    OTHER_CATEGORY_HEADER get their tagged category. A main part with no bullets
    (older free-text cells) is one claim. Drops analyzer preambles like
    'בהתבסס על המשוב, להלן הנקודות המרכזיות:'."""
    main, _, other = cell.partition(OTHER_CATEGORY_HEADER)
    bullet = r"^\s*[-*•]\s+"
    claims = [(re.sub(bullet, "", line).strip(), None)
              for line in main.splitlines() if re.match(bullet, line)]
    if not claims and main.strip():
        claims = [(main.strip(), None)]
    label_to_key = {label: key for key, label in CATEGORY_LABELS.items()}
    for line in other.splitlines():
        m = re.match(bullet + r"\[(.+?)\]\s*(.+)$", line)
        if m:
            claims.append((m.group(2).strip(), label_to_key.get(m.group(1).strip())))
    return claims


def _call_json(system_prompt: str, prompt: str, key: str | None, check=None):
    """Call Gemini in JSON mode and return parsed[key] (the whole object if key is None).
    Retries — at rising temperature, so a retry can actually differ — on invalid JSON,
    letters from foreign scripts, or a failed check(value) (returns an error string or None).
    If every attempt fails a check, returns the last parsed value for the caller to repair;
    returns None only if no attempt produced valid JSON."""
    temps = (0.0, 0.5, 0.8)
    last = None
    for attempt, temp in enumerate(temps):
        retrying = " — מנסה שוב" if attempt < len(temps) - 1 else ""
        raw = call_api(system_prompt, prompt, json_mode=True, temperature=temp)
        try:
            parsed = json.loads(raw)
            value = parsed if key is None else parsed[key]
        except (json.JSONDecodeError, KeyError, TypeError):
            print(f"    ⚠️ תשובה אינה JSON תקין{retrying}")
            continue
        last = value
        if _FOREIGN_LETTERS.search(json.dumps(value, ensure_ascii=False)):
            print(f"    ⚠️ תווים משפה זרה (למשל ערבית) בתשובה{retrying}")
            continue
        error = check(value) if check else None
        if error:
            print(f"    ⚠️ {error}{retrying}")
            continue
        return value
    return last


# A merge item is {"id", "q", "text", "sources"}: either one analyzed bullet
# (id "p3.2" = participant 3, claim 2; sources == [id]) or a claim already merged in the
# per-question summary (id "u5"; sources = the bullet ids merged into it).

def _render_bullets(items: list[dict]) -> str:
    """Raw bullets of one question, as per-participant blocks of '[id] text' lines."""
    by_participant = defaultdict(list)
    for it in items:
        by_participant[it["id"].split(".")[0]].append(it)
    return "\n\n".join(f"משתתף {p[1:]}:\n" + "\n".join(f"[{it['id']}] {it['text']}" for it in its)
                       for p, its in by_participant.items())


def _render_units(items: list[dict]) -> str:
    """Per-question merged claims, one line each, with their question and participants."""
    lines = []
    for it in items:
        participants = sorted({s.split(".")[0][1:] for s in it["sources"]}, key=int)
        lines.append(f"[{it['id']}] (שאלה: {it['q']} | משתתפים: {', '.join(participants)}) {it['text']}")
    return "\n".join(lines)


def _merge_claims(system_prompt: str, prompt: str, items: list[dict], conflict_keys) -> list[dict]:
    """Merge items via Gemini; returns [{text, sources}] with sources = bullet ids.
    Guarantees, enforced in code rather than trusted to the model:
    - nothing is silently dropped: any item the model left out is kept as its own claim;
    - no merged claim contains two items with the same conflict key (conflict_keys(item)
      returns a list of keys). Offending merges are retried, then split back into items."""
    by_id = {it["id"]: it for it in items}

    def conflicting(ids: list[str]) -> bool:
        keys = [k for i in ids if i in by_id for k in conflict_keys(by_id[i])]
        return len(keys) != len(set(keys))

    def check(claims) -> str | None:
        bad = sum(conflicting(c.get("sources", [])) for c in claims)
        return f"{bad} טענות מיזגו פריטים שאסור למזג" if bad else None

    claims = _call_json(system_prompt, prompt, "claims", check=check)
    if claims is None:
        print("    ⚠️ המיזוג נכשל — כל טענה תופיע בנפרד")
        claims = []

    merged, covered = [], set()
    for c in claims:
        ids = [i for i in c.get("sources", []) if i in by_id and i not in covered]
        text = _FOREIGN_LETTERS.sub("", str(c.get("text", ""))).strip()   # last resort
        if not ids or not text:
            continue
        covered.update(ids)
        if conflicting(ids):
            print(f"    ⚠️ פוצל מיזוג שגוי: {', '.join(ids)}")
            merged.extend({"text": by_id[i]["text"], "sources": by_id[i]["sources"]} for i in ids)
        elif len(ids) == 1:
            # Not merged with anything: keep the original text, so a model that pairs one
            # item's id with another item's text can't misattribute a claim.
            merged.append({"text": by_id[ids[0]]["text"], "sources": by_id[ids[0]]["sources"]})
        else:
            merged.append({"text": text, "sources": [s for i in ids for s in by_id[i]["sources"]]})

    missing = [it for it in items if it["id"] not in covered]
    if missing:
        print(f"    ⚠️ {len(missing)} טענות לא שויכו במיזוג — נוספו כפי שהן")
    merged.extend({"text": it["text"], "sources": it["sources"]} for it in missing)
    return merged


def _topic_groups(lines: list[str]) -> list[list[str]]:
    """Group a category's final summary bullets by related topic, for manual merge review.
    The model only returns bullet numbers; the code guarantees every bullet lands in exactly
    one group (left-out bullets become their own group). Groups are ordered by their first
    bullet's position, so they follow the summary's popularity order."""
    if len(lines) <= 1:
        return [lines] if lines else []
    prompt = "\n".join(f"{n}. {line.removeprefix('- ')}" for n, line in enumerate(lines, start=1))

    def check(groups) -> str | None:
        ok = isinstance(groups, list) and all(isinstance(g, list) for g in groups)
        return None if ok else "מבנה קבוצות שגוי"

    groups = _call_json(TOPIC_GROUPER_SYSTEM, prompt, "groups", check=check) or []
    seen, result = set(), []
    for g in groups if isinstance(groups, list) else []:
        nums = []
        for n in g if isinstance(g, list) else []:
            if isinstance(n, int) and 1 <= n <= len(lines) and n not in seen:
                seen.add(n)
                nums.append(n)
        if nums:
            result.append(sorted(nums))
    missing = [n for n in range(1, len(lines) + 1) if n not in seen]
    if missing:
        print(f"    ⚠️ {len(missing)} טענות לא שויכו לקבוצה — כל אחת בקבוצה משלה")
    result += [[n] for n in missing]
    result.sort(key=lambda g: g[0])
    return [[lines[n - 1] for n in g] for g in result]


def _participant_count(claim: dict) -> int:
    return len({s.split(".")[0] for s in claim["sources"]})


def _format_claims(claims: list[dict]) -> str:
    """Bullets ordered by unique-participant count (desc); count shown only when > 1."""
    lines = []
    for c in sorted(claims, key=_participant_count, reverse=True):
        n = _participant_count(c)
        lines.append(f"- {c['text']}" + (f" (ציינו {n} משתתפים)" if n > 1 else ""))
    return "\n".join(lines)


def build_summary_rows(all_records: list[dict]) -> list[list[str]]:
    """Build the סיכום tab rows (per-question + per-category) from Sheet 2 records.
    Gemini merges claims and tags their sources; counts and ordering are computed here."""
    # All analyzed bullets, each with a stable id p<index>.<n> so counts are by unique participant
    bullets_by_q: dict[str, list[dict]] = {q: [] for q in QUESTIONS}
    for r in all_records:
        p = str(r.get("אינדקס", "")).strip()
        n = 0
        for q in QUESTIONS:
            value = _aggregate_value(str(r.get(q, "")))
            if not value:
                continue
            for text, cat in _split_claims(value):
                n += 1
                cid = f"p{p}.{n}"
                bullets_by_q[q].append({"id": cid, "q": q, "text": text, "sources": [cid], "cat": cat})

    # ── Part 1: per-question summary ──────────────────────────────────────────
    # Never merge two bullets of the same participant: within one question they are
    # by construction different claims (the analyzer emits one bullet per claim).
    # Off-category claims (tagged by the analyzer) are left out of part 1 entirely.
    print("חלק 1: סיכום לפי שאלה")
    question_rows = [["שאלה", "סיכום משולב"]]
    units: list[dict] = []
    for q in QUESTIONS:
        items = [b for b in bullets_by_q[q] if b["cat"] is None]
        print(f"  {q[:38]}... ({len(items)} טענות)")
        if not items:
            question_rows.append([q, "אין נתונים."])
            continue
        merged = _merge_claims(MERGER_SYSTEM, f"שאלה: {q}\n\n{_render_bullets(items)}", items,
                               conflict_keys=lambda it: [it["id"].split(".")[0]])
        question_rows.append([q, _format_claims(merged)])
        units.extend({"q": q, "cat": _question_category(q), **c} for c in merged)
        print(f"    ✅ אוחדו ל-{len(merged)} טענות")

    # ── Part 2: category summary ───────────────────────────────────────────────
    # Built from the per-question claims (so every merge made in part 1 carries over) plus
    # the off-category claims, each in the category the analyzer tagged it with. Claims are
    # merged with identical ones — across questions, or within a question when part 1
    # missed a merge — but never two different claims of the same participant under the
    # same question.
    print("\nחלק 2: סיכום לפי קטגוריה")
    other = [b for q in QUESTIONS for b in bullets_by_q[q] if b["cat"] is not None]
    if other:
        print(f"  {len(other)} טענות מקטגוריה אחרת: {', '.join(b['id'] for b in other)}")
    units += [{"q": b["q"], "cat": b["cat"], "text": b["text"], "sources": b["sources"]} for b in other]
    for k, u in enumerate(units, start=1):
        u["id"] = f"u{k}"

    category_rows = [["קטגוריה", "סיכום משולב"]]
    category_lines: dict[str, list[str]] = {}
    for g in CATEGORY_GROUPS:
        items = [u for u in units if u["cat"] == g["key"]]
        print(f"  {g['name']}... ({len(items)} טענות)")
        if not items:
            category_rows.append([g["name"], "אין נתונים."])
            continue
        prompt = f"קטגוריה: {g['name']}\n\n{_render_units(items)}"
        merged = _merge_claims(CATEGORY_MERGER_SYSTEM, prompt, items,
                               conflict_keys=lambda it: [(s.split(".")[0], it["q"]) for s in it["sources"]])
        formatted = _format_claims(merged)
        category_rows.append([g["name"], formatted])
        category_lines[g["key"]] = formatted.splitlines()
        print(f"    ✅ אוחדו ל-{len(merged)} טענות")

    # ── Part 3: category bullets grouped by related topic (for manual merge decisions) ──
    print("\nחלק 3: קיבוץ לפי נושאים")
    topic_rows = [["מספר סידורי", "קטגוריה", "קבוצה"]]
    for g in CATEGORY_GROUPS:            # positive, then negative, then ideas
        lines = category_lines.get(g["key"], [])
        groups = _topic_groups(lines)
        for grp in groups:
            topic_rows.append([len(topic_rows), g["name"], "\n".join(grp)])   # serial = 1, 2, 3...
        print(f"  {g['name']}: {len(lines)} טענות ב-{len(groups)} קבוצות")

    return (
        [["", "סיכום לפי שאלה"]]
        + question_rows
        + [["", ""], ["", "סיכום לפי קטגוריה"]]
        + category_rows
        + [["", ""], ["", "קיבוץ לפי נושאים — להחלטה ידנית על מיזוגים"]]
        + topic_rows
    )


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
    all_rows = build_summary_rows(all_records)

    # ── Write to "סיכום" tab ───────────────────────────────────────────────────
    width = max(len(r) for r in all_rows)   # the topic table has 3 columns
    try:
        summary_ws = spreadsheet.worksheet("סיכום")
        summary_ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        summary_ws = spreadsheet.add_worksheet(title="סיכום", rows=len(all_rows), cols=width)
    if summary_ws.row_count < len(all_rows) or summary_ws.col_count < width:
        summary_ws.resize(rows=max(summary_ws.row_count, len(all_rows)),
                          cols=max(summary_ws.col_count, width))

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
