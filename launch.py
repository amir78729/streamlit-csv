import subprocess
import sys
import importlib.util
import threading
import webbrowser
import time
import socket

REQUIRED_PACKAGES = ['streamlit', 'pandas', 'XlsxWriter']

def install_package(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Install missing packages
for package in REQUIRED_PACKAGES:
    if importlib.util.find_spec(package) is None:
        print(f"Installing missing package: {package}")
        install_package(package)

# Choose a free port (e.g., 8501 if available)
def find_free_port(default=8501):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if s.connect_ex(('localhost', default)) != 0:
            return default
        s.bind(('', 0))
        return s.getsockname()[1]

port = find_free_port()

# Start Streamlit app in a subprocess
streamlit_process = subprocess.Popen(
    [sys.executable, "-m", "streamlit", "run", "main.py", "--server.port", str(port)],
    stdout=subprocess.PIPE,
    stderr=subprocess.STDOUT
)

# Open in browser
url = f"http://localhost:{port}"
webbrowser.open(url)

# Wait for user to press Enter to close, OR close tab
print("Streamlit app is running... Press Ctrl+C or close the tab to stop.")

try:
    streamlit_process.wait()
except KeyboardInterrupt:
    print("\nStopping Streamlit...")
    streamlit_process.terminate()
