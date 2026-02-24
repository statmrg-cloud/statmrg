# GitHub Actions로 Windows EXE 자동 빌드하기

Python 없이 실행 가능한 Windows EXE를 자동으로 만드는 방법입니다.

---

## 1회 초기 설정 (5분)

### Step 1 — GitHub 계정 준비
https://github.com 에서 무료 계정 생성 (이미 있으면 생략)

### Step 2 — 새 비공개 저장소 생성
1. GitHub 로그인 후 우상단 **"+"** → **"New repository"**
2. Repository name: `ebook-creator`
3. **Private** 선택 (소스 비공개)
4. **"Create repository"** 클릭

### Step 3 — 소스 업로드 (터미널에서 실행)

```bash
cd /Users/dowankim/ebook-creator

git init
git add .
git commit -m "초기 업로드"
git branch -M main
git remote add origin https://github.com/[내계정]/ebook-creator.git
git push -u origin main
```

> `[내계정]` 부분을 실제 GitHub 계정명으로 바꾸세요.

---

## EXE 다운로드 방법

업로드가 완료되면 GitHub Actions가 자동으로 Windows EXE를 빌드합니다 (약 5분 소요).

1. GitHub 저장소 페이지 → **"Actions"** 탭
2. **"Windows EXE 빌드"** 워크플로우 클릭
3. 완료된 실행 클릭 → 하단 **"Artifacts"** 섹션
4. **"ebook-creator-windows-exe"** 클릭해서 ZIP 다운로드

---

## 사용 방법 (Windows PC)

1. 다운로드한 ZIP 압축 해제
2. `전자책생성기.exe` 더블클릭
3. **Python 설치 불필요** — 브라우저가 자동으로 열립니다

---

## 코드 수정 후 재빌드

소스 수정 후 GitHub에 push하면 자동으로 새 EXE가 빌드됩니다:

```bash
git add .
git commit -m "수정 내용"
git push
```

---

## 수동 빌드 실행

GitHub → Actions → "Windows EXE 빌드" → **"Run workflow"** 버튼
