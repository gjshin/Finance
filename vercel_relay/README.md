# DART 중계 서버 (Vercel 서울 리전)

Streamlit Cloud 같은 **해외 서버에서는 DART에 닿지 않습니다.**
서울에 작은 중계 함수를 두고 그쪽을 거쳐 가면 해결됩니다.

```
Streamlit Cloud (해외)  →  이 함수 (서울)  →  DART
```

무료입니다. 서버를 빌리는 게 아니라 함수 하나만 올리는 방식입니다.

---

## 설치 (10분)

### 1. Vercel 가입
<https://vercel.com> — GitHub 계정으로 로그인하면 됩니다.

### 2. 프로젝트 만들기
**Add New → Project** → 이 저장소(`Finance`) 선택 →
**Root Directory** 를 `vercel_relay` 로 지정 → **Deploy**

> Root Directory를 꼭 `vercel_relay` 로 하세요. 저장소 전체가 아니라 이 폴더만 올립니다.

### 3. 리전이 서울인지 확인
**Settings → Functions → Function Region** 이 **Seoul (icn1)** 인지 봅니다.
`vercel.json` 에 이미 적어뒀지만 눈으로 확인하는 편이 좋습니다.

무료(Hobby) 플랜도 리전을 하나 고를 수 있습니다. 서울이 선택 안 되면 알려주세요.

### 4. 주소 확인
배포가 끝나면 `https://무언가.vercel.app` 주소가 나옵니다. 복사해두세요.

### 5. Streamlit Cloud에 등록
Streamlit Cloud → 앱 → **Settings → Secrets** 에 붙여넣습니다.

```toml
DART_RELAY_URL = "https://무언가.vercel.app"
```

저장하면 앱이 다시 뜹니다. 끝입니다.

---

## 잘 됐는지 확인

앱에서 조회를 한 번 돌려보세요.

- **정상 조회됨** → 완료
- **"중계 서버를 경유하도록 설정돼 있는데도 실패"** → 중계 주소나 리전을 확인
- **예전 그대로 "해외 클라우드 IP" 안내** → `DART_RELAY_URL` 이 안 읽힌 것. 오타나 저장 여부 확인

---

## 아무나 못 쓰게 막기 (선택)

중계 주소를 아는 사람은 누구나 DART 조회에 쓸 수 있습니다. 공개 데이터라 위험은 낮지만,
Vercel 무료 사용량을 남이 쓰는 게 신경 쓰이면 토큰을 거세요.

1. Vercel → **Settings → Environment Variables** 에 `RELAY_TOKEN` = 아무 긴 문자열
2. Streamlit Secrets 에도 같은 값 추가

```toml
DART_RELAY_URL = "https://무언가.vercel.app"
DART_RELAY_TOKEN = "여기에같은값"
```

설정하지 않으면 검사하지 않습니다.

---

## 알아두실 점

- **DART API 키는 중계 서버에 저장되지 않습니다.** 요청할 때마다 지나가기만 합니다.
- **DART 두 곳만 중계합니다** (`opendart.fss.or.kr`, `dart.fss.or.kr`).
  다른 주소로는 중계하지 않아 공개 프록시가 되지 않습니다.
- 회사목록(`corpCode`)은 용량이 큰 ZIP입니다. Vercel 응답 한도(4.5MB)에 걸리면
  회사 조회부터 실패할 수 있습니다. 그런 증상이 보이면 알려주세요.
- 무료 플랜에는 월 사용량 한도가 있습니다. 개인·소규모 팀 용도라면 넉넉합니다.

---

## 로컬 실행에는 영향 없음

`DART_RELAY_URL` 을 설정하지 않으면 앱은 DART로 직접 갑니다.
회계사님 PC(`run_kr.bat`)에서는 아무것도 바뀌지 않습니다.
