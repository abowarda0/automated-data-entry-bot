
# Automated Data Entry Bot for Notepad

##  Overview
This Python script automates data entry into Notepad on Windows. It fetches the first 10 blog posts from the [JSONPlaceholder API](https://jsonplaceholder.typicode.com/posts), then uses `PyAutoGUI` to type each post into Notepad. Each post is saved in a folder called `tjm-project` on the user's desktop.

---

##  Features
- Fetches posts from a public API
- Automates typing using `PyAutoGUI`
- Uses `subprocess` to launch Notepad
- Saves output as `post <id>.txt`
- Basic error handling for launch and save errors
- Optional: Uses `BotCity` and `OpenCV` for future extension (template matching, image-based automation, etc.)

---

##  Requirements

Install dependencies:
```bash
pip install -r requirements.txt
Contents of requirements.txt:

nginx
pyautogui
requests
botcity-core
botcity-maestro
opencv-python
numpy

## Demo
Click to watch: [Demo Video] https://github.com/abowarda0/automated-data-entry-bot/blob/main/demo.mp4
