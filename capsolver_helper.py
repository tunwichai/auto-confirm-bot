import requests

def solve_turnstile_captcha(api_key, sitekey, domain):
    payload = {
        "clientKey": api_key,
        "task": {
            "type": "TurnstileTaskProxyless",
            "websiteURL": domain,
            "websiteKey": sitekey
        }
    }
    resp = requests.post("https://api.capsolver.com/createTask", json=payload).json()
    task_id = resp["taskId"]
    # polling รอผลลัพธ์
    while True:
        result = requests.post("https://api.capsolver.com/getTaskResult", json={"clientKey": api_key, "taskId": task_id}).json()
        if result["status"] == "ready":
            return result["solution"]["token"]