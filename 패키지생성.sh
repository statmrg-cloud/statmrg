#!/bin/bash
# Windows 전달용 ZIP 패키지 생성 스크립트
# 실행: bash 패키지생성.sh

cd "$(dirname "$0")"
DIST="ebook-creator-windows"
ZIP_NAME="ebook-creator-windows.zip"

echo "====================================="
echo " Windows 배포 패키지 생성 중..."
echo "====================================="

# 기존 패키지 삭제
rm -rf "../$DIST" "../$ZIP_NAME"
mkdir "../$DIST"

# 필요한 파일/폴더 복사 (불필요한 것 제외)
rsync -av \
  --exclude='venv' \
  --exclude='__pycache__' \
  --exclude='*.pyc' \
  --exclude='.DS_Store' \
  --exclude='static/output/*' \
  --exclude='token_data.json' \
  --exclude='*.spec' \
  --exclude='dist' \
  --exclude='build' \
  --exclude='패키지생성.sh' \
  --exclude='.git' \
  . "../$DIST/"

# 출력 폴더 구조 유지 (빈 폴더)
mkdir -p "../$DIST/static/output"

# ZIP 생성
cd ..
zip -r "$ZIP_NAME" "$DIST" -x "*.DS_Store"
rm -rf "$DIST"

echo ""
echo "완료! → $(pwd)/$ZIP_NAME"
echo "이 ZIP 파일을 Windows PC로 복사 후 압축 해제하세요."
echo "그다음 start_windows.bat 더블클릭으로 실행하면 됩니다."
