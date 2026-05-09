#!/bin/zsh
set -e

LAUNCHER_DIR="$(cd "$(dirname "$0")" && pwd)"
SCRIPT_CANDIDATE_1="$LAUNCHER_DIR/c-d/build_full_billing_workbook.py"
SCRIPT_CANDIDATE_2="$LAUNCHER_DIR/build_full_billing_workbook.py"

if [[ -f "$SCRIPT_CANDIDATE_1" ]]; then
  SCRIPT="$SCRIPT_CANDIDATE_1"
elif [[ -f "$SCRIPT_CANDIDATE_2" ]]; then
  SCRIPT="$SCRIPT_CANDIDATE_2"
else
  echo "หา build_full_billing_workbook.py ไม่เจอ"
  echo "ให้วางไฟล์นี้ไว้ข้าง ๆ โฟลเดอร์ c-d หรือไว้โฟลเดอร์เดียวกับ script"
  if [[ -t 0 ]]; then
    read -k 1 "reply?กดปุ่มใดก็ได้เพื่อปิด..."
    echo
  fi
  exit 1
fi

BUNDLED_PY="/Users/bhusitt./.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"

if [[ -x "$BUNDLED_PY" ]]; then
  PYTHON_BIN="$BUNDLED_PY"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON_BIN="$(command -v python3)"
else
  echo "เครื่องนี้ยังไม่มี python3"
  if [[ -t 0 ]]; then
    read -k 1 "reply?กดปุ่มใดก็ได้เพื่อปิด..."
    echo
  fi
  exit 1
fi

INPUT_PATH="$1"
MONTH_ARG="$2"

if [[ -z "$INPUT_PATH" ]]; then
  echo "ลากโฟลเดอร์ c-d มาวางบนไฟล์นี้ได้เลย"
  echo -n "หรือพิมพ์ path ที่นี่: "
  read INPUT_PATH
fi

if [[ -z "$MONTH_ARG" && -t 0 ]]; then
  echo -n "ใส่เดือนข้อมูล (YYYY-MM) เช่น 2026-04 แล้วกด Enter [ค่าเริ่มต้น 2026-04]: "
  read MONTH_ARG
fi

if [[ -n "$MONTH_ARG" && ( -d "$MONTH_ARG" || "$MONTH_ARG" == */* || "$MONTH_ARG" == ~* ) ]]; then
  INPUT_PATH="$MONTH_ARG"
  MONTH_ARG=""
fi

if [[ -z "$MONTH_ARG" ]]; then
  MONTH_ARG="2026-04"
fi

if [[ -z "$INPUT_PATH" ]]; then
  echo "ยังไม่ได้ใส่ path"
  if [[ -t 0 ]]; then
    read -k 1 "reply?กดปุ่มใดก็ได้เพื่อปิด..."
    echo
  fi
  exit 1
fi

if [[ ! -d "$INPUT_PATH" ]]; then
  echo "ไม่พบโฟลเดอร์: $INPUT_PATH"
  if [[ -t 0 ]]; then
    read -k 1 "reply?กดปุ่มใดก็ได้เพื่อปิด..."
    echo
  fi
  exit 1
fi

"$PYTHON_BIN" "$SCRIPT" "$INPUT_PATH" --month "$MONTH_ARG"

echo
echo "เสร็จแล้ว"
if [[ -t 0 ]]; then
  read -k 1 "reply?กดปุ่มใดก็ได้เพื่อปิด..."
  echo
fi
