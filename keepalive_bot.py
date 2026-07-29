import time
import requests
import sys

URL = "http://localhost:5000" if len(sys.argv) < 2 else sys.argv[1]

print(f"Starting keepalive bot for {URL}")
while True:
    try:
        res = requests.get(URL)
        print(f"Pinged {URL}: {res.status_code}")
    except Exception as e:
        print(f"Failed to ping {URL}: {e}")
    time.sleep(14 * 60)
