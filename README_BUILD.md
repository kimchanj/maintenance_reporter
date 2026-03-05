# Windows 배포 빌드 방법 (PyInstaller, onedir)

## 1) 준비
- Windows 10/11
- Python 3.10~3.12 설치 (권장)
- (선택) Visual C++ Redistributable (대부분 PC에 이미 있음)

## 2) 빌드
프로젝트 루트에서 아래 중 하나 실행:

### 방법 A) 배치 파일(권장)
- build/windows/build.bat 더블클릭
- 결과: build/pyinstaller/dist/MaintenanceReporter/

### 방법 B) 직접 명령어
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconsole --clean --distpath build\pyinstaller\dist --workpath build\pyinstaller\build build\pyinstaller\maintenance_reporter.spec

## 3) 배포
- build/pyinstaller/dist/MaintenanceReporter 폴더 전체를 그대로 전달하세요.
  (폴더 안 파일들이 함께 있어야 실행됩니다)

## 4) 보안 경고(Windows SmartScreen)
- 서명 없는 exe는 경고가 뜰 수 있습니다.
- 사내 배포는 예외 등록, 외부 배포는 코드서명 인증서 사용을 고려하세요.

## 설치형 배포(권장) - Inno Setup

1) 먼저 EXE를 빌드합니다.
- `build\\windows\\build.bat`

2) 설치파일(.exe)을 만들려면 Inno Setup 6을 설치합니다.
- 설치 후 `ISCC.exe`가 PATH에 잡히면 가장 편합니다.

3) 설치파일 빌드
- `build\\windows\\build_installer.bat`

결과물 위치:
- `build\\installer_output\\MaintenanceReporter_Setup_1.0.0.exe`


## 빌드 옵션

### 1) 폴더형(안정형, onedir)
- 실행: `build\windows\build.bat`
- 결과: `build\pyinstaller\dist\MaintenanceReporter\MaintenanceReporter.exe` (폴더째 배포)

### 2) 단일 EXE(편의형, onefile)
- 실행: `build\windows\build_onefile.bat`
- 결과: `build\pyinstaller\dist_onefile\MaintenanceReporter.exe`
- 참고: 일부 PC에서 백신/SmartScreen 경고가 더 자주 뜰 수 있어요.
