# 데이터 분류 수집기 - Render 배포 가이드

## 배포 방법 (5분 소요)

### 1. GitHub에 올리기

1. https://github.com 에서 새 저장소(repository) 생성
2. 이 `deploy` 폴더의 내용을 저장소에 업로드
   - app.py
   - requirements.txt
   - render.yaml
   - static/index.html

### 2. Render에서 배포

1. https://render.com 에 접속하여 GitHub 계정으로 가입/로그인
2. 대시보드에서 **"New +"** → **"Web Service"** 클릭
3. GitHub 저장소 연결 → 방금 만든 저장소 선택
4. 설정:
   - **Name**: data-organizer (원하는 이름)
   - **Runtime**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT`
5. **"Create Web Service"** 클릭

### 3. 접속

배포 완료 후 Render가 제공하는 URL로 접속하면 됩니다.
예: `https://data-organizer-xxxx.onrender.com`

이 URL을 PC 브라우저와 안드로이드 브라우저 둘 다에서 접속하면
같은 데이터를 공유하게 됩니다.

## 주의사항

- Render 무료 플랜은 15분간 요청이 없으면 서버가 잠들어요.
  다시 접속하면 깨어나는데 30초 정도 걸릴 수 있습니다.
- 무료 플랜에서는 디스크 데이터가 재배포 시 초기화될 수 있어요.
  영구 저장이 필요하면 Render Disk($0.25/GB) 추가를 추천합니다.
