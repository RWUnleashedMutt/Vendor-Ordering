import json
import subprocess
import sys

with open('scripts.json') as f:
    scripts = json.load(f)

if len(sys.argv) < 2:
    print("Usage: python runner.py <script_name>")
    print("Available scripts:", list(scripts.keys()))
    sys.exit(1)

script_name = sys.argv[1]

if script_name not in scripts:
    print(f"Script '{script_name}' not found")
    sys.exit(1)

script_path = scripts[script_name]
python_exe = sys.executable
subprocess.run([python_exe, script_path])
