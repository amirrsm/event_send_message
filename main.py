import apprise
from apprise import NotifyFormat, NotifyType
import pandas as pd

EMAIL_HOST = "smtp.gmail.com"
EMAIL_HOST_USER = "***"
EMAIL_HOST_PASS = "***"
EMAIL_HOST_PORT = 587
EMAIL_USE_TLS = True

EVENT_TITLE = "کارگاه git"
POSTER_SRC = "https://ssces.ir:8000/media/post_git.jpg"

apr = apprise.Apprise()
df = pd.read_excel("output.xlsx")

for row in range(len(df)):
    
    apr.add(f"mailtos://{EMAIL_HOST_USER}:{EMAIL_HOST_PASS}@gmail.com:{EMAIL_HOST_PORT}?smtp={EMAIL_HOST}&to={df.loc[row, 'email']}")
        
#     msg = f"""
# <div dir="rtl">
#     <pre style="font-family: B Nazanin; font-size: 20px;">
#  سلام <strong>{df.loc[row, 'first_name']}</strong> جان
#  لطفا در روز شرکت در <u><strong>{EVENT_TITLE}</strong></u> حتما لینک کارت ورودی که برات فرستادیم همرات باشه.
#  <a href="{df.loc[row, 'ticket_link']}"> لینک کارت ورود به <strong>{EVENT_TITLE}</strong></a>
#  بزودی میبینیمت 😉</pre>
#     <img src="{POSTER_SRC}" width="50%" height="50%">
# </div>"""


    msg = f"""
<div dir="rtl">
    <pre style="font-family: B Nazanin; font-size: 16px;">
<strong>{df.loc[row, 'first_name']} {df.loc[row, 'last_name']}</strong> عزیز، سلام!

به شما خوش‌آمد می‌گوییم و ورود شما را به خانواده‌ی بزرگ دانشجویان کامپیوتر تبریک می‌گوییم. به همین مناسبت، انجمن علمی کامپیوتر قصد دارد آیین خوش‌آمدگویی ویژه‌ای را برای دانشجویان ورودی ۱۴۰۳ برگزار کند و شما را به شرکت در این برنامه دعوت می‌کنیم.

مشخصات برنامه:
تاریخ: سه‌شنبه ۱۵ آبان
ساعت: ۱۲ تا ۱۴
مکان: تالار بوزجانی

 <a href="{df.loc[row, 'ticket_link']}"> <strong>لینک تیکت</strong></a>

لطفاً از تیکت خود اسکرین شات تهیه کنید. این بارکد برای ورود شما به سالن الزامی‌ست.
در صورت نیاز به اطلاعات بیشتر، با انجمن علمی کامپیوتر در ارتباط باشید.
آیدی تلگرام: ssces_admin@
با احترام،
انجمن علمی کامپیوتر
    </pre>

</div>
"""
    title = f"آیین خوش‌آمدگویی دانشجویان ورودی 1403"
    msg_status = apr.notify(body=msg, title=title, notify_type=NotifyType.INFO, body_format=NotifyFormat.HTML)
    print(f"{'SENT' if msg_status else 'FAILED'} - {df.loc[row, 'email']}")
    apr.clear()