import subprocess

with open('uninstall_requirements.txt', 'r') as f:
    packages = f.read().splitlines()

for package in packages:
    subprocess.run(['pip', 'uninstall', '-y', package])
