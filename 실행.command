#!/bin/bash
# 두 번 눌러 실행한다. 처음 한 번은 필요한 것들을 자동으로 받는다.
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  echo "처음 실행이라 준비 중입니다. 잠시만 기다려 주세요..."
  python3 -m venv .venv
  ./.venv/bin/pip install -q --upgrade pip
  ./.venv/bin/pip install -q -r requirements.txt
fi
./.venv/bin/python parcel_app.py
