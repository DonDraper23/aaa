# Weather Mailer

`weather_mailer.py` fetches the current weather from the OpenWeatherMap API and
sends it to your email inbox every hour.

## Requirements

- Python 3
- Packages: `requests` and `schedule`

Install dependencies with:

```bash
pip install requests schedule
```

## Configuration

Set the following environment variables before running the script:

- `WEATHER_API_KEY` – API key from [OpenWeatherMap](https://openweathermap.org/)
- `WEATHER_CITY` – city to check (default is `London`)
- `SMTP_SERVER` – address of your SMTP server
- `SMTP_PORT` – SMTP port (default `587`)
- `SMTP_USER` – SMTP user name (also used as sender)
- `SMTP_PASSWORD` – SMTP password
- `RECIPIENT_EMAIL` – address to send updates to

## Usage

Run the script with Python:

```bash
python weather_mailer.py
```

It sends one update immediately and then continues running, sending a new email
once every hour. Use a process manager or cron if you want the script to run in
the background.
