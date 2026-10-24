# PyInstaller 기본 훅은 구글 API 설명 문서 605개(약 107MB)를 전부 넣는다.
# 이 앱은 Drive v3만 쓰므로 그 하나만 넣는다. (build_exe.py의 --additional-hooks-dir 로 기본 훅을 대체)
from PyInstaller.utils.hooks import collect_data_files, copy_metadata

datas = copy_metadata("google_api_python_client")
datas += collect_data_files("googleapiclient.discovery_cache", includes=["documents/drive.v3.json"])
