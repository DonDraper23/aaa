import os
import time
import requests
from email.mime.text import MIMEText
import smtplib
import schedule

API_KEY = os.environ.get('WEATHER_API_KEY')
CITY = os.environ.get('WEATHER_CITY', 'London')
SMTP_SERVER = os.environ.get('SMTP_SERVER')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '587'))
SMTP_USER = os.environ.get('SMTP_USER')
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD')
RECIPIENT = os.environ.get('RECIPIENT_EMAIL')
SENDER = SMTP_USER


def get_weather() -> str:
    if not API_KEY:
        raise ValueError('WEATHER_API_KEY environment variable not set')
    resp = requests.get(
        'https://api.openweathermap.org/data/2.5/weather',
        params={'q': CITY, 'appid': API_KEY, 'units': 'metric'},
        timeout=10,
    )
    resp.raise_for_status()
    data = resp.json()
    desc = data['weather'][0]['description']
    temp = data['main']['temp']
    return f'Weather in {CITY}: {desc}, {temp} °C'


def send_email(text: str) -> None:
    if not (SMTP_SERVER and SMTP_USER and SMTP_PASSWORD and RECIPIENT):
        raise ValueError('SMTP configuration environment variables missing')
    msg = MIMEText(text)
    msg['Subject'] = f'Weather update for {CITY}'
    msg['From'] = SENDER
    msg['To'] = RECIPIENT

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg)


def job() -> None:
    text = get_weather()
    send_email(text)


def main() -> None:
    job()  # run once at startup
    schedule.every().hour.do(job)
    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == '__main__':
    main()
