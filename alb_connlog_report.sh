#!/usr/bin/env bash
#===============================================================================
# alb_connlog_report.sh
#
#   AWS Application Load Balancer (ALB) の接続ログ (connection log) を解析し、
#   Excel (xlsx) 形式のレポートを出力する。
#
#   S3 に出力された接続ログ (conn_log.*.log.gz) を入力ディレクトリに配置し、
#   そのディレクトリを指定して実行する。サブディレクトリ配下の *.gz も対象とする。
#
# 動作環境:
#   RHEL 9.x (bash 5.1 / GNU coreutils / gzip / gawk / python3)
#   xlsx の生成には OS 標準の python3 の標準ライブラリのみを使用するため、
#   openpyxl などの追加パッケージやインターネット接続は不要。
#
# 使い方:
#   ./alb_connlog_report.sh [オプション] <入力ディレクトリ>   (詳細は -h を参照)
#
# 出力シート:
#   サマリー             : 処理概要、TLS 検証結果・リスナーポート・TLS プロトコル等の集計
#   接続ログ明細         : 接続ログ 1 行 = 1 行の明細 (見出し・フィルター・ウィンドウ枠固定付き)
#   時間帯別件数         : 1 時間単位の接続件数 (成功 / 失敗)
#   クライアントIP別件数 : クライアント IP ごとの接続件数、最初 / 最後の接続日時
#   ファイル別件数       : 入力ファイルごとのレコード数・読込結果
#   解析エラー           : 解析できなかった行 (発生時のみ)
#   項目説明             : 各列の説明、tls_verify_status のエラーコード一覧
#
# 終了コード:
#   0: 正常終了 / 1: 引数・実行環境エラー / 2: 入力ファイルなし / 3: レポート作成失敗
#===============================================================================
set -euo pipefail

readonly SCRIPT_NAME="${0##*/}"
SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
readonly SCRIPT_DIR

#--- 既定値 --------------------------------------------------------------------
PYTHON_BIN="${PYTHON_BIN:-python3}"   # 使用する python (環境変数 PYTHON_BIN で変更可)
OUTPUT_FILE=""                        # 出力ファイル (未指定時は SCRIPT_DIR/output/ 配下)
TZ_OFFSET="+09:00"                    # 表示用タイムゾーン (UTC からの時差)
MAX_ROWS=1000000                      # 1 シートあたりの最大明細行数
TOP_N=20                              # サマリーのランキング表示件数
SORT_BY_TIME=1                        # 1: timestamp 昇順に並び替える / 0: ファイル順

#--- 実行時に設定する変数 ------------------------------------------------------
INPUT_DIR=""
TZ_MINUTES=0
TZ_LABEL=""
WORK_DIR=""
MANIFEST=""
PY_SCRIPT=""
TMP_OUTPUT=""
GZ_FILES=()

#===============================================================================
# 共通関数
#===============================================================================
log()   { printf '%s [%s] %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$1" "$2" >&2; }
info()  { log INFO "$*"; }
warn()  { log WARN "$*"; }
error() { log ERROR "$*"; }
die()   { error "$1"; exit "${2:-1}"; }

cleanup() {
    if [[ -n $WORK_DIR && -d $WORK_DIR ]]; then
        rm -rf -- "$WORK_DIR"
    fi
    if [[ -n $TMP_OUTPUT && -f $TMP_OUTPUT ]]; then
        rm -f -- "$TMP_OUTPUT"
    fi
}

usage() {
    cat <<EOF
使い方: ${SCRIPT_NAME} [オプション] <入力ディレクトリ>

  ALB 接続ログ (conn_log.*.log.gz) を解析し、Excel (xlsx) レポートを出力します。
  入力ディレクトリ配下 (サブディレクトリを含む) の *.gz ファイルをすべて対象とします。

オプション:
  -o <出力ファイル>  出力する xlsx ファイルのパス (ディレクトリを指定した場合はその配下に出力)
                     既定: <スクリプト配置ディレクトリ>/output/alb_connlog_report_YYYYMMDD_HHMMSS.xlsx
  -z <時差>          表示用タイムゾーンの UTC からの時差 (例: +09:00, +9, 0, -05:00)
                     既定: +09:00 (日本時間)
  -m <行数>          1 シートあたりの最大明細行数 (1～1048575)。超えた分は次のシートに出力
                     既定: ${MAX_ROWS}
  -t <件数>          サマリーのランキング (クライアント IP 等) の表示件数。既定: ${TOP_N}
  -n                 timestamp での並び替えを行わない (ファイル名順・ファイル内の出現順で出力)
  -h                 このヘルプを表示

環境変数:
  PYTHON_BIN         使用する python コマンド (既定: python3)
  TMPDIR             作業用ディレクトリの作成先 (並び替えの一時ファイルを作成。既定: /tmp)

実行例:
  ${SCRIPT_NAME} ./input
  ${SCRIPT_NAME} -o /var/tmp/report.xlsx -z 0 ./input
EOF
}

#===============================================================================
# 引数・実行環境のチェック
#===============================================================================
parse_args() {
    local opt
    while getopts ":o:z:m:t:nh" opt; do
        case $opt in
            o) OUTPUT_FILE=$OPTARG ;;
            z) TZ_OFFSET=$OPTARG ;;
            m) MAX_ROWS=$OPTARG ;;
            t) TOP_N=$OPTARG ;;
            n) SORT_BY_TIME=0 ;;
            h) usage; exit 0 ;;
            :) die "オプション -$OPTARG には値を指定してください (詳細は -h を参照)" ;;
            \?) die "不明なオプションです: -$OPTARG (詳細は -h を参照)" ;;
        esac
    done
    shift $((OPTIND - 1))

    if (( $# != 1 )); then
        usage >&2
        exit 1
    fi
    INPUT_DIR=$1
}

validate_options() {
    local sign="+" hh mm

    # 表示用タイムゾーン (+09:00 / +9 / 0 / -05:30 などを受け付ける)
    if [[ $TZ_OFFSET =~ ^([+-]?)([0-9]{1,2})(:?([0-9]{2}))?$ ]]; then
        [[ -n ${BASH_REMATCH[1]} ]] && sign=${BASH_REMATCH[1]}
        hh=$((10#${BASH_REMATCH[2]}))
        mm=$((10#${BASH_REMATCH[4]:-0}))
    else
        die "-z の値が不正です: ${TZ_OFFSET} (例: +09:00, +9, 0)"
    fi
    if (( hh > 14 || mm > 59 )); then
        die "-z の値が範囲外です: ${TZ_OFFSET}"
    fi
    TZ_MINUTES=$((hh * 60 + mm))
    if [[ $sign == "-" ]]; then
        TZ_MINUTES=$((-TZ_MINUTES))
    fi
    if (( TZ_MINUTES == 0 )); then
        TZ_LABEL="UTC"
    else
        TZ_LABEL=$(printf 'UTC%s%02d:%02d' "$sign" "$hh" "$mm")
    fi

    if [[ ! $MAX_ROWS =~ ^[0-9]{1,7}$ ]] || (( 10#$MAX_ROWS < 1 || 10#$MAX_ROWS > 1048575 )); then
        die "-m の値が不正です: ${MAX_ROWS} (1～1048575 を指定してください)"
    fi
    MAX_ROWS=$((10#$MAX_ROWS))

    if [[ ! $TOP_N =~ ^[0-9]{1,6}$ ]] || (( 10#$TOP_N < 1 )); then
        die "-t の値が不正です: ${TOP_N} (1 以上の整数を指定してください)"
    fi
    TOP_N=$((10#$TOP_N))
}

check_prerequisites() {
    local cmd

    if (( BASH_VERSINFO[0] < 4 || (BASH_VERSINFO[0] == 4 && BASH_VERSINFO[1] < 4) )); then
        die "bash 4.4 以上が必要です (現在: ${BASH_VERSION})"
    fi
    for cmd in find sort gzip awk stat du cut mktemp; do
        command -v "$cmd" >/dev/null 2>&1 || die "必要なコマンドが見つかりません: ${cmd}"
    done
    if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
        die "python が見つかりません: ${PYTHON_BIN} (RHEL の場合は 'dnf install python3' でインストールしてください)"
    fi
    if ! "$PYTHON_BIN" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 6) else 1)'; then
        die "python 3.6 以上が必要です: ${PYTHON_BIN}"
    fi
}

#===============================================================================
# 入出力の準備
#===============================================================================
prepare_input() {
    [[ -e $INPUT_DIR ]] || die "入力ディレクトリが存在しません: ${INPUT_DIR}"
    [[ -d $INPUT_DIR ]] || die "入力パスがディレクトリではありません: ${INPUT_DIR}"
    [[ -r $INPUT_DIR && -x $INPUT_DIR ]] || die "入力ディレクトリを読み取れません: ${INPUT_DIR}"
    INPUT_DIR=$(cd -- "$INPUT_DIR" && pwd)

    # サブディレクトリも含めて *.gz を収集し、パス名順に並べる
    mapfile -d '' -t GZ_FILES < <(find -L "$INPUT_DIR" -type f -name '*.gz' -print0 | LC_ALL=C sort -z)
    if (( ${#GZ_FILES[@]} == 0 )); then
        die "入力ディレクトリに gz ファイルがありません: ${INPUT_DIR}" 2
    fi
}

prepare_output() {
    local default_name out_dir
    default_name="alb_connlog_report_$(date '+%Y%m%d_%H%M%S').xlsx"

    if [[ -z $OUTPUT_FILE ]]; then
        OUTPUT_FILE="${SCRIPT_DIR}/output/${default_name}"
    elif [[ -d $OUTPUT_FILE || $OUTPUT_FILE == */ ]]; then
        OUTPUT_FILE="${OUTPUT_FILE%/}/${default_name}"
    fi
    if [[ ${OUTPUT_FILE,,} != *.xlsx ]]; then
        OUTPUT_FILE+=".xlsx"
    fi

    out_dir=$(dirname -- "$OUTPUT_FILE")
    mkdir -p -- "$out_dir" || die "出力ディレクトリを作成できません: ${out_dir}"
    [[ -w $out_dir ]] || die "出力ディレクトリに書き込めません: ${out_dir}"
    out_dir=$(cd -- "$out_dir" && pwd)
    OUTPUT_FILE="${out_dir}/${OUTPUT_FILE##*/}"
    if [[ -e $OUTPUT_FILE ]]; then
        warn "出力ファイルが既に存在するため上書きします: ${OUTPUT_FILE}"
    fi

    # 書き込み途中のファイルが残らないよう、一時ファイルに出力してから置き換える
    TMP_OUTPUT="${out_dir}/.${OUTPUT_FILE##*/}.$$.tmp"
}

#===============================================================================
# ログの展開
#   各 gz ファイルを展開し「ソースファイル<TAB>行番号<TAB>ログ行」の形式で
#   標準出力へ出力する。ファイルごとの読込結果はマニフェスト (TSV) に記録する。
#===============================================================================
extract_logs() {
    local f rel size status msg
    local -a rcs

    for f in "${GZ_FILES[@]}"; do
        rel=${f#"$INPUT_DIR"/}
        size=$(stat -c %s -- "$f" 2>/dev/null || echo 0)

        set +e
        gzip -dc -- "$f" 2>"${WORK_DIR}/gzip.err" \
            | LC_ALL=C SRC_FILE=$rel awk '
                BEGIN { OFS = "\t" }
                { sub(/\r$/, "") }
                $0 != "" { print ENVIRON["SRC_FILE"], FNR, $0 }'
        rcs=("${PIPESTATUS[@]}")
        set -e

        # 後続の処理が異常終了した (SIGPIPE 等) 場合は展開を打ち切る
        if (( rcs[1] > 128 )); then
            return 1
        fi

        # gzip のメッセージは先頭に空行が入ることがあるため、最初の空でない行を使う
        msg=$(LC_ALL=C awk 'NF { gsub(/[\t\r]/, " "); print; exit }' "${WORK_DIR}/gzip.err")
        msg=${msg//"$f"/"$rel"}
        if (( rcs[0] == 0 && rcs[1] == 0 )); then
            status="OK"
        elif (( rcs[0] == 2 && rcs[1] == 0 )); then
            status="WARN"   # gzip の警告 (trailing garbage ignored 等)。展開結果は利用する
            warn "gz ファイルの展開で警告が発生しました: ${rel} : ${msg}"
        else
            status="NG"
            [[ -n $msg ]] || msg="展開に失敗しました (gzip 終了コード: ${rcs[0]}, awk 終了コード: ${rcs[1]})"
            warn "gz ファイルを読み込めませんでした (読み込めた行のみ出力します): ${rel} : ${msg}"
        fi
        printf '%s\t%s\t%s\t%s\n' "$rel" "$size" "$status" "$msg" >> "$MANIFEST"
    done
}

run_pipeline() {
    local sort_mode="time"
    (( SORT_BY_TIME )) || sort_mode="file"

    local -a py_args=(
        --manifest "$MANIFEST"
        --output "$TMP_OUTPUT"
        --input-dir "$INPUT_DIR"
        "--tz-minutes=${TZ_MINUTES}"
        --tz-label "$TZ_LABEL"
        --max-rows "$MAX_ROWS"
        --top-n "$TOP_N"
        --sort-mode "$sort_mode"
    )

    if (( SORT_BY_TIME )); then
        # 行の先頭が timestamp (ISO 8601 / UTC) のため、ログ行 (3 列目) の文字列順 = 時刻順
        extract_logs \
            | LC_ALL=C sort -t $'\t' -k3,3 -s -T "$WORK_DIR" \
            | "$PYTHON_BIN" -I "$PY_SCRIPT" "${py_args[@]}"
    else
        extract_logs | "$PYTHON_BIN" -I "$PY_SCRIPT" "${py_args[@]}"
    fi
}

#===============================================================================
# xlsx 生成処理 (python3 標準ライブラリのみ使用)
#===============================================================================
write_python_script() {
    cat > "$1" <<'PYEOF'
# -*- coding: utf-8 -*-
"""ALB 接続ログを解析し、xlsx レポートを出力する (alb_connlog_report.sh から呼び出される)。

標準入力: 「ソースファイル<TAB>行番号<TAB>ログ行」形式の行
Python 3.6 以降の標準ライブラリのみで動作する。
"""
import argparse
import datetime
import math
import re
import sys
import time
import zipfile
from collections import Counter

#--- 定数 ----------------------------------------------------------------------
EXCEL_MAX_ROWS = 1048576                 # Excel の最大行数
EXCEL_MAX_CELL_CHARS = 32767             # Excel の 1 セルあたりの最大文字数
SHEET_XML_LIMIT = 1800 * 1024 * 1024     # 1 シートの XML サイズ上限 (ZIP64 を使わないよう 2GiB 未満に抑える)
MAX_STORED_ERRORS = 100000               # 解析エラーシートに出力する最大行数
PROGRESS_INTERVAL = 200000               # 進捗表示の間隔 (行)
EXCEL_EPOCH = datetime.datetime(1899, 12, 30)

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

DETAIL_SHEET_NAME = '接続ログ明細'
REFERENCE_URL = 'https://docs.aws.amazon.com/ja_jp/elasticloadbalancing/latest/application/load-balancer-connection-logs.html'

TOKEN_RE = re.compile(r'"(?:[^"\\]|\\.)*"|\S+')
TS_RE = re.compile(r'^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?Z$')
VALIDITY_RE = re.compile(r'^NotBefore=([^;]*);NotAfter=(.*)$', re.IGNORECASE)
INT_RE = re.compile(r'^\d{1,15}$', re.ASCII)
ILLEGAL_XML_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

# ALB 接続ログのフィールド数 (conn_trace_id が追加される前の形式は 11 個)
MIN_FIELDS = 11

# tls_verify_status のエラーコード (AWS ドキュメント「接続ログのエントリ」より)
ERROR_CODES = [
    ('ClientCertMaxChainDepthExceeded', 'Max chain depth exceeded',
     '証明書チェーンの深さが上限を超えています'),
    ('ClientCertMaxSizeExceeded', 'Max certificate size exceeded',
     'クライアント証明書のサイズが上限を超えています'),
    ('ClientCertCrlHit', 'Client certificate was revoked by CA',
     'クライアント証明書が CA によって失効されています (CRL に該当)'),
    ('ClientCertCrlProcessingError', 'CRL processing error',
     'CRL (証明書失効リスト) の処理でエラーが発生しました'),
    ('ClientCertUntrusted', 'Client certificate is untrusted',
     'クライアント証明書が信頼されていません'),
    ('ClientCertNotYetValid', 'Client certificate is not yet valid',
     'クライアント証明書の有効期間が開始されていません'),
    ('ClientCertExpired', 'Client certificate has expired',
     'クライアント証明書の有効期限が切れています'),
    ('ClientCertTypeUnsupported', 'Client certificate type unsupported',
     'サポートされていない種類のクライアント証明書です'),
    ('ClientCertInvalid', 'Client certificate is invalid',
     'クライアント証明書が無効です'),
    ('ClientCertRejected', 'Client certificate is rejected by custom server validation',
     'カスタムサーバー検証によってクライアント証明書が拒否されました'),
    ('UnmappedConnectionError', 'Unmapped runtime connection error',
     '分類されていない実行時の接続エラーです'),
]
ERROR_CODE_JA = dict((code, ja) for code, _, ja in ERROR_CODES)

#--- セルスタイル (styles.xml の cellXfs の並びと一致させること) ---------------
(ST_DEFAULT, ST_HEAD_LOG, ST_HEAD_ADD,
 ST_TEXT, ST_INT, ST_DEC3, ST_DATETIME,
 ST_TEXT_NG, ST_INT_NG, ST_DEC3_NG, ST_DATETIME_NG,
 ST_TITLE, ST_SECTION, ST_TABLE_HEAD, ST_LABEL, ST_COUNT, ST_PERCENT,
 ST_WRAP, ST_HOUR, ST_NOTE, ST_COUNT_NG, ST_LEGEND_LOG, ST_LEGEND_ADD,
 ST_KV_COUNT, ST_KV_DATETIME, ST_KV_DATETIME_S, ST_CENTER) = range(27)
NG_OFFSET = ST_TEXT_NG - ST_TEXT         # 通常スタイル → 失敗行スタイルへの差分

# numFmtId 164 以降はユーザー定義書式
NUM_FMTS = [
    (164, 'yyyy/mm/dd hh:mm:ss.000'),
    (165, '0.000'),
    (166, '0.0%'),
    (167, 'yyyy/mm/dd hh:mm'),
    (168, 'yyyy/mm/dd hh:mm:ss'),
]
FONT = 'Meiryo UI'
# (サイズ, 太字, 色)
FONTS = [
    (10, False, None),          # 0 標準
    (10, True, 'FFFFFFFF'),     # 1 見出し
    (14, True, 'FF1F4E78'),     # 2 タイトル
    (10, True, None),           # 3 太字
    (10, False, 'FF9C0006'),    # 4 失敗行
    (11, True, 'FF1F4E78'),     # 5 セクション見出し
    (9, False, 'FF595959'),     # 6 注記
]
# 塗りつぶし色 (0, 1 は Excel の予約: none / gray125)
FILLS = [None, None, 'FF1F4E78', 'FF548235', 'FFFFC7CE', 'FFDDEBF7', 'FFF2F2F2']
ALIGNMENTS = {
    'center_wrap': '<alignment horizontal="center" vertical="center" wrapText="1"/>',
    'center': '<alignment horizontal="center" vertical="center"/>',
    'left': '<alignment horizontal="left" vertical="center"/>',
    'wrap': '<alignment vertical="center" wrapText="1"/>',
    'vcenter': '<alignment vertical="center"/>',
}
# (numFmtId, fontId, fillId, borderId, 配置)
CELL_XFS = [
    (0, 0, 0, 0, None),             # ST_DEFAULT
    (0, 1, 2, 1, 'center_wrap'),    # ST_HEAD_LOG     見出し (ALB ログ項目)
    (0, 1, 3, 1, 'center_wrap'),    # ST_HEAD_ADD     見出し (付加情報)
    (0, 0, 0, 1, 'vcenter'),        # ST_TEXT
    (1, 0, 0, 1, 'vcenter'),        # ST_INT
    (165, 0, 0, 1, 'vcenter'),      # ST_DEC3
    (164, 0, 0, 1, 'vcenter'),      # ST_DATETIME
    (0, 4, 4, 1, 'vcenter'),        # ST_TEXT_NG
    (1, 4, 4, 1, 'vcenter'),        # ST_INT_NG
    (165, 4, 4, 1, 'vcenter'),      # ST_DEC3_NG
    (164, 4, 4, 1, 'vcenter'),      # ST_DATETIME_NG
    (0, 2, 0, 0, None),             # ST_TITLE
    (0, 5, 0, 0, None),             # ST_SECTION
    (0, 3, 5, 1, 'center_wrap'),    # ST_TABLE_HEAD
    (0, 3, 6, 1, 'vcenter'),        # ST_LABEL
    (3, 0, 0, 1, 'vcenter'),        # ST_COUNT        #,##0
    (166, 0, 0, 1, 'vcenter'),      # ST_PERCENT
    (0, 0, 0, 1, 'wrap'),           # ST_WRAP
    (167, 0, 0, 1, 'vcenter'),      # ST_HOUR
    (0, 6, 0, 0, None),             # ST_NOTE
    (3, 4, 4, 1, 'vcenter'),        # ST_COUNT_NG
    (0, 1, 2, 1, 'center'),         # ST_LEGEND_LOG
    (0, 1, 3, 1, 'center'),         # ST_LEGEND_ADD
    (3, 0, 0, 1, 'left'),           # ST_KV_COUNT
    (164, 0, 0, 1, 'left'),         # ST_KV_DATETIME
    (168, 0, 0, 1, 'left'),         # ST_KV_DATETIME_S
    (0, 0, 0, 1, 'center'),         # ST_CENTER
]

#--- 明細シートの列定義 --------------------------------------------------------
LOG, ADD = 'ALBログ項目', '付加情報'
# (見出し(日本語), 項目名, 区分, 列幅, 数値のスタイル, 説明)  ※ {tz} は表示用タイムゾーンに置換
COLUMNS = [
    ('No.', '', ADD, 8, ST_INT,
     '明細の通し番号 (明細シートが複数に分かれる場合も通し番号)。'),
    ('接続日時({tz})', 'timestamp ※変換', ADD, 24, ST_DATETIME,
     'timestamp を表示用タイムゾーン ({tz}) に変換した日時。Excel の日付値のため、並べ替えや日付フィルターが利用できる。'),
    ('接続日時(UTC)', 'timestamp', LOG, 28, None,
     'ロードバランサーがクライアントとの接続の確立に成功した、または失敗した時刻 (ISO 8601 形式、UTC)。'),
    ('クライアントIP', 'client_ip', LOG, 17, None,
     '接続元クライアントの IP アドレス。'),
    ('クライアントポート', 'client_port', LOG, 15, ST_INT,
     '接続元クライアントのポート番号。'),
    ('リスナーポート', 'listener_port', LOG, 14, ST_INT,
     'クライアントからの接続を受け付けたロードバランサーのリスナーのポート番号。'),
    ('TLSプロトコル', 'tls_protocol', LOG, 13, None,
     '[HTTPS リスナー] ハンドシェイクで使用された SSL/TLS プロトコル。SSL/TLS 以外の接続の場合は "-"。'),
    ('TLS暗号スイート', 'tls_cipher', LOG, 30, None,
     '[HTTPS リスナー] 使用された暗号スイート (OpenSSL 形式の名称)。SSL/TLS 以外の接続の場合は "-"。'),
    ('TLSハンドシェイク時間(秒)', 'tls_handshake_latency', LOG, 23, ST_DEC3,
     '[HTTPS リスナー] ハンドシェイクの確立に要した合計時間 (秒、ミリ秒精度)。SSL/TLS 以外の接続、またはハンドシェイクに失敗した場合は "-"。'),
    ('クライアント証明書 サブジェクト', 'leaf_client_cert_subject', LOG, 44, None,
     '[HTTPS リスナー] リーフクライアント証明書のサブジェクト名。SSL/TLS 以外の接続、mTLS が有効でない場合、証明書を解析できない場合は "-"。'),
    ('証明書 有効期間開始', 'leaf_client_cert_validity (NotBefore)', LOG, 27, None,
     '[HTTPS リスナー] リーフクライアント証明書の有効期間の開始日時 (leaf_client_cert_validity の NotBefore 部分、ISO 8601 形式)。該当しない場合は "-"。'),
    ('証明書 有効期間終了', 'leaf_client_cert_validity (NotAfter)', LOG, 27, None,
     '[HTTPS リスナー] リーフクライアント証明書の有効期間の終了日時 (leaf_client_cert_validity の NotAfter 部分、ISO 8601 形式)。該当しない場合は "-"。'),
    ('証明書 シリアル番号', 'leaf_client_cert_serial_number', LOG, 34, None,
     '[HTTPS リスナー] リーフクライアント証明書のシリアル番号。該当しない場合は "-"。'),
    ('TLS検証ステータス', 'tls_verify_status', LOG, 30, None,
     '[HTTPS リスナー] 接続要求のステータス。接続に成功した場合は "Success"、失敗した場合は "Failed:<エラーコード>"。'),
    ('検証結果の説明', 'tls_verify_status ※補足', ADD, 44, None,
     'tls_verify_status の内容を日本語で説明したもの (エラーコードの説明は AWS ドキュメントの定義に基づく)。'),
    ('接続トレースID', 'conn_trace_id', LOG, 26, None,
     '接続ごとに付与される一意の追跡 ID。この接続で送信されたリクエストのアクセスログにも同じ ID が記録されるため、アクセスログとの突合に利用できる。'),
    ('追加フィールド', '※未定義の項目', ADD, 16, None,
     '上記以外にログ行に含まれていたフィールド (AWS がログ形式に項目を追加した場合に格納)。通常は空欄。'),
    ('ソースファイル', '※入力ファイル', ADD, 60, None,
     'このレコードを読み込んだ gz ファイル (入力ディレクトリからの相対パス)。'),
    ('行番号', '※ファイル内', ADD, 10, ST_INT,
     'ソースファイル内 (展開後) の行番号。'),
]


#===============================================================================
# 共通関数
#===============================================================================
def log(level, msg):
    line = '%s [%s] %s\n' % (datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), level, msg)
    sys.stderr.buffer.write(line.encode('utf-8', 'replace'))
    sys.stderr.flush()


def col_letter(index):
    """0 始まりの列番号を A, B, ..., AA 形式に変換する。"""
    letters = ''
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


COL_LETTERS = [col_letter(i) for i in range(64)]


def esc(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def esc_attr(text):
    return esc(text).replace('"', '&quot;')


def cell_xml(ref, value, style):
    """1 セル分の XML を返す。文字列はインライン文字列として出力する。"""
    cls = value.__class__
    if cls is int:
        return '<c r="%s" s="%d"><v>%d</v></c>' % (ref, style, value)
    if cls is float:
        if math.isfinite(value):
            return '<c r="%s" s="%d"><v>%r</v></c>' % (ref, style, value)
        value = str(value)
    elif value is None:
        return '<c r="%s" s="%d"/>' % (ref, style)
    elif cls is not str:
        value = str(value)
    if not value:
        return '<c r="%s" s="%d"/>' % (ref, style)
    if len(value) > EXCEL_MAX_CELL_CHARS:
        value = value[:EXCEL_MAX_CELL_CHARS]
    value = ILLEGAL_XML_RE.sub('', value)
    space = ''
    if value[:1].isspace() or value[-1:].isspace() or '\n' in value:
        space = ' xml:space="preserve"'
    return '<c r="%s" s="%d" t="inlineStr"><is><t%s>%s</t></is></c>' % (ref, style, space, esc(value))


def excel_serial(dt):
    """datetime を Excel のシリアル値 (1900 年日付システム) に変換する。"""
    delta = dt - EXCEL_EPOCH
    return delta.days + (delta.seconds + delta.microseconds / 1000000.0) / 86400.0


def to_int(value):
    return int(value) if INT_RE.match(value) else value


def to_float(value):
    if value == '-':
        return value
    try:
        number = float(value)
    except ValueError:
        return value
    return number if math.isfinite(number) else value


def unquote(token):
    if len(token) >= 2 and token[0] == '"' and token[-1] == '"':
        return token[1:-1].replace('\\"', '"')
    return token


def text_width(text):
    """Excel の列幅単位でのおおよその文字幅 (全角は半角の 1.4 倍程度)。"""
    return sum(1.4 if ord(ch) > 0x7f else 1.0 for ch in text)


def wrapped_height(text, width, line_height=14.0):
    """折り返し表示に必要な行の高さ (1 行に収まる場合は None)。"""
    lines = 0
    for part in text.split('\n'):
        lines += max(1, int(math.ceil(text_width(part) / float(width))))
    return None if lines <= 1 else line_height * lines + 4


def describe_status(status):
    if status == 'Success':
        return '接続成功'
    if status == '-':
        return '該当なし'
    if status.startswith('Failed'):
        code = status.partition(':')[2]
        return '接続失敗: ' + ERROR_CODE_JA.get(code, '未定義のエラーコード (AWS ドキュメントを参照)')
    return ''


#===============================================================================
# xlsx 出力 (ワークシートはストリーミングで ZIP に書き込む)
#===============================================================================
def zip_info(name):
    info = zipfile.ZipInfo(name, date_time=time.localtime()[:6])
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    return info


class Sheet(object):
    """1 枚のワークシート。行を順に書き込み、close() で XML を完成させる。"""

    def __init__(self, zf, part_no, name, widths, freeze_rows=0, freeze_cols=0, selected=False):
        self.name = name
        self.part = 'xl/worksheets/sheet%d.xml' % part_no
        self.ncols = len(widths)
        self.rows = 0
        self.size = 0
        self.autofilter = False
        self.filter_ref = None
        self.merges = []
        self._buf = []
        self._buf_len = 0
        self._fh = zf.open(zip_info(self.part), 'w')

        pane = ''
        if freeze_rows or freeze_cols:
            top_left = col_letter(freeze_cols) + str(freeze_rows + 1)
            if freeze_rows and freeze_cols:
                active = 'bottomRight'
            elif freeze_rows:
                active = 'bottomLeft'
            else:
                active = 'topRight'
            split = ''
            if freeze_cols:
                split += ' xSplit="%d"' % freeze_cols
            if freeze_rows:
                split += ' ySplit="%d"' % freeze_rows
            pane = ('<pane%s topLeftCell="%s" activePane="%s" state="frozen"/>'
                    '<selection pane="%s" activeCell="%s" sqref="%s"/>'
                    % (split, top_left, active, active, top_left, top_left))
        cols = ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i, i, w)
                       for i, w in enumerate(widths, 1))
        self._write('%s<worksheet xmlns="%s" xmlns:r="%s"><sheetViews>'
                    '<sheetView workbookViewId="0"%s>%s</sheetView></sheetViews>'
                    '<sheetFormatPr defaultRowHeight="15"/><cols>%s</cols><sheetData>'
                    % (XML_DECL, NS_MAIN, NS_REL, ' tabSelected="1"' if selected else '', pane, cols))

    def _write(self, text):
        self._buf.append(text)
        self._buf_len += len(text)
        if self._buf_len >= 1 << 20:
            self._flush()

    def _flush(self):
        if self._buf:
            data = ''.join(self._buf).encode('utf-8')
            self._fh.write(data)
            self.size += len(data)
            self._buf = []
            self._buf_len = 0

    def write_row(self, cells=(), height=None):
        """cells: (値, スタイル) のリスト。None の要素は空セルとして出力しない。"""
        self.rows += 1
        r = self.rows
        rs = str(r)
        out = ['<row r="%d" ht="%s" customHeight="1">' % (r, height) if height else '<row r="%d">' % r]
        for i, cell in enumerate(cells):
            if cell is not None:
                out.append(cell_xml(COL_LETTERS[i] + rs, cell[0], cell[1]))
        out.append('</row>')
        self._write(''.join(out))
        return r

    def write_values(self, values, num_styles, text_style, offset):
        """明細行用の高速版。文字列は制御文字を除去済みであること。
        数値は num_styles、文字列は text_style のスタイルに offset を加えて出力する。"""
        self.rows += 1
        rs = str(self.rows)
        out = ['<row r="%s">' % rs]
        append = out.append
        for col, value, num_style in zip(COL_LETTERS, values, num_styles):
            cls = value.__class__
            if cls is str:
                style = text_style + offset
                if not value:
                    append('<c r="%s%s" s="%d"/>' % (col, rs, style))
                    continue
                if len(value) > EXCEL_MAX_CELL_CHARS:
                    value = value[:EXCEL_MAX_CELL_CHARS]
                if '&' in value or '<' in value or '>' in value:
                    value = esc(value)
                append('<c r="%s%s" s="%d" t="inlineStr"><is><t>%s</t></is></c>' % (col, rs, style, value))
            elif cls is int:
                append('<c r="%s%s" s="%d"><v>%d</v></c>' % (col, rs, num_style + offset, value))
            elif cls is float:
                append('<c r="%s%s" s="%d"><v>%r</v></c>' % (col, rs, num_style + offset, value))
            else:
                append(cell_xml(col + rs, value, text_style + offset))
        append('</row>')
        self._write(''.join(out))

    def merge(self, ref):
        self.merges.append(ref)

    def close(self):
        last_col = col_letter(self.ncols - 1)
        tail = ['</sheetData>']
        if self.autofilter and self.rows:
            tail.append('<autoFilter ref="A1:%s%d"/>' % (last_col, self.rows))
            self.filter_ref = '$A$1:$%s$%d' % (last_col, self.rows)
        if self.merges:
            tail.append('<mergeCells count="%d">%s</mergeCells>'
                        % (len(self.merges), ''.join('<mergeCell ref="%s"/>' % m for m in self.merges)))
        tail.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        if self.rows:
            # 「数値が文字列として保存されています」の警告表示を抑止する
            tail.append('<ignoredErrors><ignoredError sqref="A1:%s%d" numberStoredAsText="1"/></ignoredErrors>'
                        % (last_col, self.rows))
        tail.append('</worksheet>')
        self._write(''.join(tail))
        self._flush()
        self._fh.close()


class Workbook(object):
    def __init__(self, path):
        self._zf = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)
        self._sheets = []

    def add_sheet(self, name, widths, **kwargs):
        sheet = Sheet(self._zf, len(self._sheets) + 1, name, widths, **kwargs)
        self._sheets.append(sheet)
        return sheet

    def close(self, order, title):
        """order: 表示順に並べたシートのリスト (すべて close 済みであること)。"""
        rid = dict((s.part, 'rId%d' % i) for i, s in enumerate(self._sheets, 1))
        sheets_xml = ''.join('<sheet name="%s" sheetId="%d" r:id="%s"/>' % (esc_attr(s.name), i, rid[s.part])
                             for i, s in enumerate(order, 1))
        names_xml = ''.join('<definedName name="_xlnm._FilterDatabase" localSheetId="%d" hidden="1">%s</definedName>'
                            % (i, esc("'%s'!%s" % (s.name.replace("'", "''"), s.filter_ref)))
                            for i, s in enumerate(order) if s.filter_ref)
        if names_xml:
            names_xml = '<definedNames>%s</definedNames>' % names_xml
        self._writestr('xl/workbook.xml',
                       '%s<workbook xmlns="%s" xmlns:r="%s"><bookViews><workbookView activeTab="0"/></bookViews>'
                       '<sheets>%s</sheets>%s</workbook>' % (XML_DECL, NS_MAIN, NS_REL, sheets_xml, names_xml))

        rels = ''.join('<Relationship Id="%s" Type="%s/worksheet" Target="worksheets/%s"/>'
                       % (rid[s.part], NS_REL, s.part.split('/')[-1]) for s in self._sheets)
        rels += '<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/>' % (len(self._sheets) + 1, NS_REL)
        self._writestr('xl/_rels/workbook.xml.rels',
                       '%s<Relationships xmlns="%s">%s</Relationships>' % (XML_DECL, NS_PKG_REL, rels))
        self._writestr('xl/styles.xml', build_styles())

        self._writestr('_rels/.rels',
                       '%s<Relationships xmlns="%s">'
                       '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>'
                       '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                       '<Relationship Id="rId3" Type="%s/extended-properties" Target="docProps/app.xml"/>'
                       '</Relationships>' % (XML_DECL, NS_PKG_REL, NS_REL, NS_REL))
        now = time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime())
        self._writestr('docProps/core.xml',
                       '%s<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
                       ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
                       ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                       '<dc:title>%s</dc:title><dc:creator>alb_connlog_report.sh</dc:creator>'
                       '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
                       '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
                       '</cp:coreProperties>' % (XML_DECL, esc(title), now, now))
        self._writestr('docProps/app.xml',
                       '%s<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                       '<Application>alb_connlog_report.sh</Application></Properties>' % XML_DECL)

        ct = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
        overrides = ''.join('<Override PartName="/%s" ContentType="%s.worksheet+xml"/>' % (s.part, ct)
                            for s in self._sheets)
        self._writestr('[Content_Types].xml',
                       '%s<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                       '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                       '<Default Extension="xml" ContentType="application/xml"/>'
                       '<Override PartName="/xl/workbook.xml" ContentType="%s.sheet.main+xml"/>'
                       '<Override PartName="/xl/styles.xml" ContentType="%s.styles+xml"/>'
                       '%s'
                       '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                       '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                       '</Types>' % (XML_DECL, ct, ct, overrides))
        self._zf.close()

    def _writestr(self, name, text):
        self._zf.writestr(zip_info(name), text.encode('utf-8'))


def build_styles():
    num_fmts = ''.join('<numFmt numFmtId="%d" formatCode="%s"/>' % (i, esc_attr(code)) for i, code in NUM_FMTS)
    fonts = ''
    for size, bold, color in FONTS:
        fonts += '<font>%s<sz val="%d"/>%s<name val="%s"/><family val="3"/><charset val="128"/></font>' % (
            '<b/>' if bold else '', size, '<color rgb="%s"/>' % color if color else '', FONT)
    fills = '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    for color in FILLS[2:]:
        fills += '<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>' % color
    side = '<%s style="thin"><color rgb="FFBFBFBF"/></%s>'
    borders = ('<border><left/><right/><top/><bottom/><diagonal/></border><border>%s<diagonal/></border>'
               % ''.join(side % (s, s) for s in ('left', 'right', 'top', 'bottom')))
    xfs = ''
    for num_fmt, font, fill, border, align in CELL_XFS:
        attrs = ('numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"'
                 ' applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"' % (num_fmt, font, fill, border))
        if align:
            xfs += '<xf %s applyAlignment="1">%s</xf>' % (attrs, ALIGNMENTS[align])
        else:
            xfs += '<xf %s/>' % attrs
    return ('%s<styleSheet xmlns="%s">'
            '<numFmts count="%d">%s</numFmts>'
            '<fonts count="%d">%s</fonts>'
            '<fills count="%d">%s</fills>'
            '<borders count="2">%s</borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="%d">%s</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>'
            % (XML_DECL, NS_MAIN, len(NUM_FMTS), num_fmts, len(FONTS), fonts, len(FILLS), fills,
               borders, len(CELL_XFS), xfs))


#===============================================================================
# 解析
#===============================================================================
class Stats(object):
    """集計値を保持する。"""

    def __init__(self):
        self.records = 0
        self.success = 0
        self.failed = 0
        self.status = Counter()
        self.listener_port = Counter()
        self.protocol = Counter()
        self.cipher = Counter()
        self.ips = {}          # client_ip -> [件数, 成功, 失敗, 最初の日時, 最後の日時]
        self.subjects = {}     # 証明書サブジェクト -> [件数, 失敗]
        self.hourly = {}       # 1 時間単位の日時 -> [件数, 成功, 失敗]
        self.latency_ms = Counter()   # ハンドシェイク時間 (ミリ秒単位) の度数
        self.latency_sum = 0.0
        self.first = None
        self.last = None
        self.extra_records = 0
        self.file_records = Counter()
        self.file_errors = Counter()
        self.errors = []
        self.error_count = 0

    def add_error(self, src, line_no, reason, text):
        self.error_count += 1
        self.file_errors[src] += 1
        if len(self.errors) < MAX_STORED_ERRORS:
            self.errors.append((src, line_no, reason, text))

    def add_record(self, src, fields, local_dt, serial, status, latency):
        self.records += 1
        self.file_records[src] += 1
        ok = status == 'Success'
        ng = status.startswith('Failed')
        if ok:
            self.success += 1
        elif ng:
            self.failed += 1
        self.status[status] += 1
        self.listener_port[fields[3]] += 1
        self.protocol[fields[4]] += 1
        self.cipher[fields[5]] += 1
        if len(fields) > 12:
            self.extra_records += 1

        ip = self.ips.get(fields[1])
        if ip is None:
            ip = self.ips[fields[1]] = [0, 0, 0, serial, serial]
        ip[0] += 1
        if ok:
            ip[1] += 1
        elif ng:
            ip[2] += 1
        if serial is not None:
            if ip[3] is None or serial < ip[3]:
                ip[3] = serial
            if ip[4] is None or serial > ip[4]:
                ip[4] = serial

        subject = fields[7]
        if subject != '-':
            sub = self.subjects.get(subject)
            if sub is None:
                sub = self.subjects[subject] = [0, 0]
            sub[0] += 1
            if ng:
                sub[1] += 1

        if local_dt is not None:
            key = local_dt.replace(minute=0, second=0, microsecond=0)
            hour = self.hourly.get(key)
            if hour is None:
                hour = self.hourly[key] = [0, 0, 0]
            hour[0] += 1
            if ok:
                hour[1] += 1
            elif ng:
                hour[2] += 1
            if self.first is None or serial < self.first:
                self.first = serial
            if self.last is None or serial > self.last:
                self.last = serial

        if latency.__class__ is float:
            self.latency_ms[int(round(latency * 1000))] += 1
            self.latency_sum += latency

    def latency_percentile(self, pct):
        n = sum(self.latency_ms.values())
        target = max(1, int(math.ceil(n * pct / 100.0)))
        acc = 0
        for ms in sorted(self.latency_ms):
            acc += self.latency_ms[ms]
            if acc >= target:
                return ms / 1000.0
        return None


def parse_log_line(text):
    """ログ行をフィールドに分解する。戻り値: (フィールドのリスト, timestamp の一致結果, エラー理由)"""
    tokens = TOKEN_RE.findall(text)
    if len(tokens) < MIN_FIELDS:
        return None, None, 'フィールド数が不足しています (%d 個 / 必要数 %d 個以上)' % (len(tokens), MIN_FIELDS)
    match = TS_RE.match(tokens[0])
    if match is None:
        return None, None, '先頭のフィールドが timestamp (ISO 8601 形式) ではありません'
    return [unquote(t) for t in tokens], match, None


def to_local_datetime(match, tz_delta):
    year, month, day, hour, minute, second, frac = match.groups()
    micro = int((frac or '0')[:6].ljust(6, '0'))
    try:
        dt = datetime.datetime(int(year), int(month), int(day), int(hour), int(minute), int(second), micro)
        return dt + tz_delta
    except (ValueError, OverflowError):
        return None


def split_validity(value):
    if value == '-':
        return '-', '-'
    m = VALIDITY_RE.match(value)
    if m:
        return m.group(1), m.group(2)
    return value, ''


class DetailWriter(object):
    """明細シートの出力。最大行数を超えたら次のシートに切り替える。"""

    def __init__(self, book, tz_label, max_rows):
        self.book = book
        self.max_rows = max_rows
        self.sheets = []
        self.sheet = None
        self.header = []
        for ja, en, kind, _, _, _ in COLUMNS:
            label = ja.format(tz=tz_label) + ('\n' + en if en else '')
            self.header.append((label, ST_HEAD_LOG if kind == LOG else ST_HEAD_ADD))
        self.widths = [c[3] for c in COLUMNS]
        self.num_styles = [c[4] or ST_TEXT for c in COLUMNS]

    def _next_sheet(self):
        if self.sheet is not None:
            self.sheet.close()
        self.sheet = self.book.add_sheet(DETAIL_SHEET_NAME, self.widths, freeze_rows=1, freeze_cols=2)
        self.sheet.autofilter = True
        self.sheet.write_row(self.header, height=46)
        self.sheets.append(self.sheet)

    def write(self, values, failed):
        sheet = self.sheet
        if sheet is None or sheet.rows > self.max_rows or sheet.size > SHEET_XML_LIMIT:
            self._next_sheet()
            sheet = self.sheet
        sheet.write_values(values, self.num_styles, ST_TEXT, NG_OFFSET if failed else 0)

    def finish(self):
        if self.sheet is None:
            self._next_sheet()
        self.sheet.close()
        if len(self.sheets) > 1:
            for i, sheet in enumerate(self.sheets, 1):
                sheet.name = '%s_%d' % (DETAIL_SHEET_NAME, i)
        return self.sheets


#===============================================================================
# 集計シートの出力
#===============================================================================
def pct(count, total):
    return count / float(total) if total else 0.0


def write_summary(book, st, args, detail_sheets, files):
    widths = [46, 14, 10, 16, 50]
    sh = book.add_sheet('サマリー', widths, selected=True)
    total = st.records
    tz = args.tz_label

    def section(title):
        sh.write_row()
        sh.write_row([(title, ST_SECTION)], height=21)

    def kv(label, value, style=ST_TEXT):
        r = sh.write_row([(label, ST_LABEL), (value, style), ('', style), ('', style), ('', style)])
        sh.merge('B%d:E%d' % (r, r))

    def note(text):
        sh.write_row([(text, ST_NOTE)])

    def table_head(*labels):
        sh.write_row([(label, ST_TABLE_HEAD) for label in labels], height=20)

    def no_data(ncols):
        sh.write_row([('(データなし)', ST_TEXT)] + [('', ST_TEXT)] * (ncols - 1))

    def counter_table(title, key_label, counter):
        section(title)
        table_head(key_label, '件数', '割合')
        if not counter:
            no_data(3)
        for key, count in sorted(counter.items(), key=lambda kv_: (-kv_[1], kv_[0])):
            sh.write_row([(key, ST_TEXT), (count, ST_COUNT), (pct(count, total), ST_PERCENT)])

    sh.write_row([('ALB 接続ログ 解析レポート', ST_TITLE)], height=26)

    # ■ 処理概要
    file_ng = sum(1 for f in files if f[2] == 'NG')
    section('■ 処理概要')
    kv('レポート作成日時', excel_serial(datetime.datetime.now()), ST_KV_DATETIME_S)
    kv('入力ディレクトリ', args.input_dir)
    kv('対象ファイル数 (*.gz)', len(files), ST_KV_COUNT)
    kv('読込エラーのファイル数', file_ng, ST_KV_COUNT)
    kv('総レコード数 (明細行数)', total, ST_KV_COUNT)
    kv('TLS 検証 成功 (Success) 件数', st.success, ST_KV_COUNT)
    kv('TLS 検証 失敗 (Failed) 件数', st.failed, ST_KV_COUNT)
    kv('その他 ("-" 等) 件数', total - st.success - st.failed, ST_KV_COUNT)
    kv('解析エラー行数', st.error_count, ST_KV_COUNT)
    kv('ユニーククライアント IP 数', len(st.ips), ST_KV_COUNT)
    kv('最初の接続日時 (%s)' % tz, st.first if st.first is not None else '-', ST_KV_DATETIME)
    kv('最後の接続日時 (%s)' % tz, st.last if st.last is not None else '-', ST_KV_DATETIME)
    kv('表示用タイムゾーン', tz)
    kv('明細の並び順', 'timestamp の昇順' if args.sort_mode == 'time' else '入力ファイル名順 (ファイル内は出現順)')
    kv('明細シート', '、'.join(s.name for s in detail_sheets) + ' (%d シート)' % len(detail_sheets))
    if file_ng:
        note('※ 読込エラーのファイルがあります。「ファイル別件数」シートを確認してください。')
    if st.error_count:
        note('※ 解析できなかった行があります。「解析エラー」シートを確認してください。')
        if st.error_count > len(st.errors):
            note('※ 解析エラーが多いため、「解析エラー」シートには先頭 %d 行のみ出力しています。' % len(st.errors))
    if st.extra_records:
        note('※ 未定義のフィールドを含む行が %d 行あります (明細の「追加フィールド」列に出力)。' % st.extra_records)

    # ■ TLS 検証結果
    section('■ TLS 検証結果別の件数 (tls_verify_status)')
    table_head('tls_verify_status', '件数', '割合', '説明', '')
    sh.merge('D%d:E%d' % (sh.rows, sh.rows))
    if not st.status:
        no_data(5)
    for status, count in sorted(st.status.items(), key=lambda kv_: (-kv_[1], kv_[0])):
        off = NG_OFFSET if status.startswith('Failed') else 0
        r = sh.write_row([(status, ST_TEXT + off), (count, ST_COUNT_NG if off else ST_COUNT),
                          (pct(count, total), ST_PERCENT), (describe_status(status), ST_TEXT + off),
                          ('', ST_TEXT + off)])
        sh.merge('D%d:E%d' % (r, r))

    counter_table('■ リスナーポート別の件数 (listener_port)', 'listener_port', st.listener_port)
    counter_table('■ TLS プロトコル別の件数 (tls_protocol)', 'tls_protocol', st.protocol)
    counter_table('■ TLS 暗号スイート別の件数 (tls_cipher)', 'tls_cipher', st.cipher)

    # ■ TLS ハンドシェイク時間
    section('■ TLS ハンドシェイク時間 (tls_handshake_latency)')
    table_head('統計項目', '値', '単位')
    n = sum(st.latency_ms.values())
    if n:
        rows = [('対象件数 ("-" を除く)', n, ST_COUNT, '件'),
                ('最小', min(st.latency_ms) / 1000.0, ST_DEC3, '秒'),
                ('平均', st.latency_sum / n, ST_DEC3, '秒'),
                ('中央値', st.latency_percentile(50), ST_DEC3, '秒'),
                ('90 パーセンタイル', st.latency_percentile(90), ST_DEC3, '秒'),
                ('95 パーセンタイル', st.latency_percentile(95), ST_DEC3, '秒'),
                ('99 パーセンタイル', st.latency_percentile(99), ST_DEC3, '秒'),
                ('最大', max(st.latency_ms) / 1000.0, ST_DEC3, '秒')]
        for label, value, style, unit in rows:
            sh.write_row([(label, ST_TEXT), (value, style), (unit, ST_CENTER)])
    else:
        no_data(3)

    # ■ クライアント IP 上位
    section('■ クライアント IP 別の件数 上位 %d 件 (client_ip)' % args.top_n)
    table_head('client_ip', '件数', '割合', 'うち失敗(Failed)')
    ranked = sorted(st.ips.items(), key=lambda kv_: (-kv_[1][0], kv_[0]))[:args.top_n]
    if not ranked:
        no_data(4)
    for ip, v in ranked:
        sh.write_row([(ip, ST_TEXT), (v[0], ST_COUNT), (pct(v[0], total), ST_PERCENT),
                      (v[2], ST_COUNT_NG if v[2] else ST_COUNT)])
    note('※ 全 %d 件のクライアント IP は「クライアントIP別件数」シートを参照してください。' % len(st.ips))

    # ■ クライアント証明書 上位
    section('■ クライアント証明書のサブジェクト別の件数 上位 %d 件 (leaf_client_cert_subject)' % args.top_n)
    table_head('leaf_client_cert_subject', '件数', '割合', 'うち失敗(Failed)')
    ranked = sorted(st.subjects.items(), key=lambda kv_: (-kv_[1][0], kv_[0]))[:args.top_n]
    if not ranked:
        no_data(4)
    for subject, v in ranked:
        sh.write_row([(subject, ST_WRAP), (v[0], ST_COUNT), (pct(v[0], total), ST_PERCENT),
                      (v[1], ST_COUNT_NG if v[1] else ST_COUNT)], height=wrapped_height(subject, widths[0]))
    note('※ "-" (mTLS 未使用など証明書情報なし) は集計対象外です。')
    sh.close()
    return sh


def write_hourly_sheet(book, st, tz):
    sh = book.add_sheet('時間帯別件数', [22, 13, 15, 15, 13], freeze_rows=1)
    sh.autofilter = True
    sh.write_row([(h, ST_TABLE_HEAD) for h in ('日時 (1 時間単位)\n%s' % tz, '接続件数', '成功\n(Success)',
                                               '失敗\n(Failed)', 'その他')], height=32)
    for key in sorted(st.hourly):
        total, ok, ng = st.hourly[key]
        sh.write_row([(excel_serial(key), ST_HOUR), (total, ST_COUNT), (ok, ST_COUNT),
                      (ng, ST_COUNT_NG if ng else ST_COUNT), (total - ok - ng, ST_COUNT)])
    sh.close()
    return sh


def write_ip_sheet(book, st, tz):
    sh = book.add_sheet('クライアントIP別件数', [8, 18, 12, 14, 14, 10, 24, 24], freeze_rows=1)
    sh.autofilter = True
    sh.write_row([(h, ST_TABLE_HEAD) for h in (
        'No.', 'クライアントIP\nclient_ip', '接続件数', '成功\n(Success)', '失敗\n(Failed)', 'その他',
        '最初の接続日時\n%s' % tz, '最後の接続日時\n%s' % tz)], height=32)
    ranked = sorted(st.ips.items(), key=lambda kv_: (-kv_[1][0], kv_[0]))
    if len(ranked) > EXCEL_MAX_ROWS - 1:
        log('WARN', 'クライアント IP が多いため、上位 %d 件のみ出力します' % (EXCEL_MAX_ROWS - 1))
        ranked = ranked[:EXCEL_MAX_ROWS - 1]
    for i, (ip, v) in enumerate(ranked, 1):
        total, ok, ng, first, last = v
        sh.write_row([(i, ST_INT), (ip, ST_TEXT), (total, ST_COUNT), (ok, ST_COUNT),
                      (ng, ST_COUNT_NG if ng else ST_COUNT), (total - ok - ng, ST_COUNT),
                      (first, ST_DATETIME), (last, ST_DATETIME)])
    sh.close()
    return sh


def write_file_sheet(book, st, files):
    sh = book.add_sheet('ファイル別件数', [7, 90, 15, 13, 13, 12, 60], freeze_rows=1)
    sh.autofilter = True
    sh.write_row([(h, ST_TABLE_HEAD) for h in (
        'No.', 'ファイル名 (入力ディレクトリからの相対パス)', 'サイズ (バイト)', 'レコード数',
        '解析エラー行数', '読込結果', 'メッセージ')], height=32)
    labels = {'OK': '正常', 'WARN': '警告', 'NG': '読込エラー'}
    for i, (name, size, status, message) in enumerate(files, 1):
        off = NG_OFFSET if status == 'NG' else 0
        count_style = ST_COUNT_NG if off else ST_COUNT
        sh.write_row([(i, ST_INT + off), (name, ST_TEXT + off), (to_int(size), count_style),
                      (st.file_records.get(name, 0), count_style), (st.file_errors.get(name, 0), count_style),
                      (labels.get(status, status), ST_TEXT + off), (message, ST_TEXT + off)])
    sh.close()
    return sh


def write_error_sheet(book, st):
    sh = book.add_sheet('解析エラー', [8, 60, 10, 52, 120], freeze_rows=1)
    sh.autofilter = True
    sh.write_row([(h, ST_TABLE_HEAD) for h in (
        'No.', 'ソースファイル', '行番号', 'エラー内容', 'ログ行 (原文)')], height=32)
    for i, (src, line_no, reason, text) in enumerate(st.errors, 1):
        sh.write_row([(i, ST_INT), (src, ST_TEXT), (line_no, ST_INT), (reason, ST_TEXT), (text, ST_TEXT)])
    sh.close()
    return sh


def write_field_sheet(book, tz):
    widths = [7, 32, 52, 14, 110]
    sh = book.add_sheet('項目説明', widths, freeze_rows=1)
    sh.write_row([(h, ST_TABLE_HEAD) for h in ('列', '見出し', '項目名 (ALB ログ)', '区分', '説明')], height=22)
    for i, (ja, en, kind, _, _, desc) in enumerate(COLUMNS):
        desc = desc.format(tz=tz)
        sh.write_row([(col_letter(i), ST_CENTER), (ja.format(tz=tz), ST_TEXT), (en, ST_TEXT),
                      (kind, ST_LEGEND_LOG if kind == LOG else ST_LEGEND_ADD), (desc, ST_WRAP)],
                     height=wrapped_height(desc, widths[4]))

    sh.write_row()
    sh.write_row([('■ 補足', ST_SECTION)], height=21)
    for text in (
            '・明細シートの見出しは 1 行目が日本語の項目名、2 行目が ALB ログの項目名です。'
            '見出しの色は、青 = ALB ログの項目、緑 = 本ツールで付加した情報 を表します。',
            '・値が "-" の項目は「該当なし」を表します (SSL/TLS 以外の接続、mTLS が有効でない場合など)。',
            '・tls_verify_status が "Failed" の行は、明細シート上で赤色の背景で強調表示しています。',
            '・明細シートは見出し行にフィルターを設定し、見出し行と左端 2 列 (No.、接続日時) を固定表示しています。',
            '・明細の行数が 1 シートの上限を超える場合は「%s_1」「%s_2」… のように複数シートに分割して出力します。'
            % (DETAIL_SHEET_NAME, DETAIL_SHEET_NAME),
            '・参考: AWS ドキュメント「Application Load Balancer の接続ログ」 %s' % REFERENCE_URL):
        sh.write_row([(text, ST_NOTE)])

    sh.write_row()
    sh.write_row([('■ tls_verify_status のエラーコード一覧 (Failed:<エラーコード>)', ST_SECTION)], height=21)
    sh.write_row([None, ('エラーコード', ST_TABLE_HEAD), ('説明 (AWS ドキュメント)', ST_TABLE_HEAD),
                  ('説明', ST_TABLE_HEAD), ('', ST_TABLE_HEAD)])
    sh.merge('D%d:E%d' % (sh.rows, sh.rows))
    for code, en, ja in ERROR_CODES:
        r = sh.write_row([None, (code, ST_TEXT), (en, ST_TEXT), (ja, ST_TEXT), ('', ST_TEXT)])
        sh.merge('D%d:E%d' % (r, r))
    sh.close()
    return sh


#===============================================================================
# メイン処理
#===============================================================================
def read_manifest(path):
    files = []
    with open(path, 'rb') as fh:
        for raw in fh:
            parts = raw.decode('utf-8', 'replace').rstrip('\r\n').split('\t')
            if len(parts) >= 3:
                files.append((parts[0], parts[1], parts[2], parts[3] if len(parts) > 3 else ''))
    return files


def parse_args():
    p = argparse.ArgumentParser(description='ALB 接続ログを解析し xlsx レポートを出力する')
    p.add_argument('--manifest', required=True, help='入力ファイルの一覧 (TSV)')
    p.add_argument('--output', required=True, help='出力する xlsx ファイル')
    p.add_argument('--input-dir', default='', help='入力ディレクトリ (表示用)')
    p.add_argument('--tz-minutes', type=int, default=540, help='表示用タイムゾーンの UTC からの時差 (分)')
    p.add_argument('--tz-label', default='UTC+09:00', help='表示用タイムゾーンの表記')
    p.add_argument('--max-rows', type=int, default=1000000, help='1 シートあたりの最大明細行数')
    p.add_argument('--top-n', type=int, default=20, help='ランキングの表示件数')
    p.add_argument('--sort-mode', choices=('time', 'file'), default='time', help='明細の並び順')
    args = p.parse_args()
    args.max_rows = max(1, min(args.max_rows, EXCEL_MAX_ROWS - 1))
    return args


def main():
    args = parse_args()
    tz_delta = datetime.timedelta(minutes=args.tz_minutes)
    st = Stats()
    book = Workbook(args.output)
    detail = DetailWriter(book, args.tz_label, args.max_rows)
    desc_cache = {}

    log('INFO', 'ログの解析と明細の出力を開始します')
    for raw in sys.stdin.buffer:
        parts = raw.decode('utf-8', 'replace').rstrip('\r\n').split('\t', 2)
        if len(parts) != 3:
            continue
        src, line_no, text = parts
        line_no = to_int(line_no)
        text = ILLEGAL_XML_RE.sub('', text)

        fields, match, reason = parse_log_line(text)
        if fields is None:
            st.add_error(src, line_no, reason, text)
            continue

        local_dt = to_local_datetime(match, tz_delta)
        serial = excel_serial(local_dt) if local_dt is not None else None
        not_before, not_after = split_validity(fields[8])
        status = fields[10]
        desc = desc_cache.get(status)
        if desc is None:
            desc = desc_cache[status] = describe_status(status)
        latency = to_float(fields[6])

        st.add_record(src, fields, local_dt, serial, status, latency)
        detail.write([st.records, serial, fields[0], fields[1], to_int(fields[2]), to_int(fields[3]),
                      fields[4], fields[5], latency, fields[7], not_before, not_after, fields[9],
                      status, desc, fields[11] if len(fields) > 11 else '', ' '.join(fields[12:]),
                      src, line_no],
                     status.startswith('Failed'))
        if st.records % PROGRESS_INTERVAL == 0:
            log('INFO', '  明細を出力中: %s 行' % format(st.records, ','))

    detail_sheets = detail.finish()
    log('INFO', '明細の出力が完了しました: %s 行 (%d シート)' % (format(st.records, ','), len(detail_sheets)))

    files = read_manifest(args.manifest)
    log('INFO', '集計シートを出力します')
    order = [write_summary(book, st, args, detail_sheets, files)] + detail_sheets
    order.append(write_hourly_sheet(book, st, args.tz_label))
    order.append(write_ip_sheet(book, st, args.tz_label))
    order.append(write_file_sheet(book, st, files))
    if st.errors:
        order.append(write_error_sheet(book, st))
    order.append(write_field_sheet(book, args.tz_label))
    book.close(order, 'ALB 接続ログ 解析レポート')

    log('INFO', '集計結果: レコード %s 件 (成功 %s / 失敗 %s) / 解析エラー %s 行 / 対象ファイル %d 件'
        % (format(st.records, ','), format(st.success, ','), format(st.failed, ','),
           format(st.error_count, ','), len(files)))
    if st.records == 0:
        log('WARN', '有効な接続ログのレコードがありませんでした')
    if st.error_count:
        log('WARN', '解析できなかった行が %s 行あります (「解析エラー」シートを参照)' % format(st.error_count, ','))


if __name__ == '__main__':
    main()
PYEOF
}

#===============================================================================
# メイン
#===============================================================================
main() {
    local start=$SECONDS size

    if [[ ${1:-} == "--help" ]]; then
        usage
        exit 0
    fi
    parse_args "$@"
    validate_options
    check_prerequisites
    prepare_input
    prepare_output

    trap cleanup EXIT
    trap 'exit 130' INT
    trap 'exit 143' TERM
    WORK_DIR=$(mktemp -d "${TMPDIR:-/tmp}/alb_connlog_report.XXXXXX") \
        || die "作業ディレクトリを作成できません (TMPDIR=${TMPDIR:-/tmp})"
    MANIFEST="${WORK_DIR}/manifest.tsv"
    PY_SCRIPT="${WORK_DIR}/alb_connlog_to_xlsx.py"
    : > "$MANIFEST"
    write_python_script "$PY_SCRIPT"

    info "入力ディレクトリ: ${INPUT_DIR} (gz ファイル ${#GZ_FILES[@]} 件)"
    info "出力ファイル    : ${OUTPUT_FILE}"
    info "表示タイムゾーン: ${TZ_LABEL} / 並び順: $( (( SORT_BY_TIME )) && echo 'timestamp 昇順' || echo 'ファイル順')"

    if ! run_pipeline; then
        die "レポートの作成に失敗しました" 3
    fi
    mv -f -- "$TMP_OUTPUT" "$OUTPUT_FILE"
    TMP_OUTPUT=""

    size=$(du -h -- "$OUTPUT_FILE" | cut -f1)
    info "レポートを出力しました: ${OUTPUT_FILE} (${size}, 処理時間 $((SECONDS - start)) 秒)"
}

main "$@"
