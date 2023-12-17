# oleobj 툴로 directory 내부의 모든 파일 파싱

import os
import subprocess

directory = "./malole"

# 디렉토리 순회
for filename in os.listdir(directory):
    if filename.endswith(".doc") or filename.endswith(".docx"):
        filepath = os.path.join(directory, filename)
        print(f"Processing: {filepath}")

        # oleobj 명령어 실행
        result = subprocess.run(["oleobj", "-i", filepath], capture_output=True, text=True)

        # 결과 출력
        print(result.stdout)
        if result.stderr:
            print("Error:", result.stderr)

print("OLE extraction complete.")