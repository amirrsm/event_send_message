import apprise
from apprise import NotifyFormat, NotifyType
import pandas as pd

EMAIL_HOST = "smtp.gmail.com"
EMAIL_HOST_USER = "amirrezasalarimanesh1302"
EMAIL_HOST_PASS = "ndravdxxtgrvvbki"
EMAIL_HOST_PORT = 587
EMAIL_USE_TLS = True

EVENT_TITLE = "کارگاه git"
POSTER_SRC = "https://ssces.ir:8000/media/post_git.jpg"

apr = apprise.Apprise()
df = pd.read_excel("output.xlsx")

for row in range(len(df)):
    
    apr.add(f"mailtos://{EMAIL_HOST_USER}:{EMAIL_HOST_PASS}@gmail.com:{EMAIL_HOST_PORT}?smtp={EMAIL_HOST}&to={df.loc[row, 'email']}")
        
    msg = f"""
<div dir="rtl">
    <pre style="font-family: B Nazanin; font-size: 20px;">
 سلام <strong>{df.loc[row, 'first_name']}</strong> جان
 لطفا در روز شرکت در <u><strong>{EVENT_TITLE}</strong></u> حتما لینک کارت ورودی که برات فرستادیم همرات باشه.
 <a href="{df.loc[row, 'ticket_link']}"> لینک کارت ورود به <strong>{EVENT_TITLE}</strong></a>
 بزودی میبینیمت 😉</pre>
    <img src="{POSTER_SRC}" width="50%" height="50%">
</div>"""
    title = f"کارت ورود به {EVENT_TITLE} "
    msg_status = apr.notify(body=msg, title=title, notify_type=NotifyType.INFO, body_format=NotifyFormat.HTML)
    print(f"{'SENT' if msg_status else 'FAILED'} - {df.loc[row, 'email']}")
    apr.clear()