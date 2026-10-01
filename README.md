# 쿠팡 광고 보고서 생성기

쿠팡 Wing 광고 보고서 엑셀(직전월·금월)을 올리면 계정별 HTML 보고서를 만들고, 여러 계정 보고서를 하나로 통합합니다.
생성된 HTML은 다운로드 후에도 `✏️ 편집` 버튼으로 수정·저장할 수 있습니다.

## 온라인 배포 (Streamlit Community Cloud · 무료)

> 광고 성과는 영업 자료이므로 **비공개(Private) 저장소 + 비공개 앱**으로 운영합니다.
> 무료 플랜은 비공개 앱을 1개까지 만들 수 있습니다.

### 1. 깃허브에 올리기
1. <https://github.com/new> 에서 저장소 생성 — 이름 예: `coupang-ad-report`, **Private** 선택, README 추가 체크 해제
2. 이 폴더를 그 저장소로 푸시 (처음 한 번 깃허브 로그인 창이 뜹니다)
   ```
   git remote add origin https://github.com/<내아이디>/coupang-ad-report.git
   git push -u origin main
   ```

### 2. 앱 배포
1. <https://share.streamlit.io> 에서 **Continue with GitHub**로 로그인 → 비공개 저장소 접근 허용
2. **Create app → Deploy a public app from GitHub** 대신 저장소 목록에서 위 저장소 선택
   - Branch: `main` · Main file path: `app.py`
   - App URL: 원하는 주소 입력 (예: `coupang-ad-report`)
3. **Deploy** → 2~3분 뒤 `https://<주소>.streamlit.app` 링크 생성

### 3. 접속자 지정
앱 우측 상단 **Share** → 이메일로 팀원 초대. 초대된 사람만 접속할 수 있습니다.

### 수정 반영
코드를 고친 뒤 `git push` 하면 앱이 자동으로 다시 배포됩니다.

## 참고
- 한동안 접속이 없으면 앱이 잠들며, 다음 접속 시 깨우는 데 수십 초가 걸립니다.
- 업로드한 엑셀은 앱 서버 메모리에서만 처리되고 저장소에는 남지 않습니다. 엑셀·보고서·설정 JSON은 `.gitignore`로 커밋에서 제외됩니다.
- 로컬에서 실행하려면 `실행.bat`을 더블클릭하세요.
- 테스트: `py -m unittest discover -s tests`
