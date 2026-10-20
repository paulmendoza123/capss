#!/usr/bin/env bash
# Creates the HTTPS certificates for SPARK (run ONCE on the Raspberry Pi).
#
#   bash deploy/make_cert.sh <PI-IP-ADDRESS> [output-folder]
#   e.g. bash deploy/make_cert.sh 192.168.1.50
#
# It makes two things:
#   1. A small private "certificate authority" (spark-ca.crt / spark-ca.key).
#      Install spark-ca.crt on a device ONLY if you want that device to show
#      no warning at all. Keep spark-ca.key secret (never share or commit it).
#   2. The server certificate nginx uses (spark.crt / spark.key), valid for
#      https://spark.local , https://spark , and https://<PI-IP-ADDRESS>.
#
# Valid for 397 days. If it expires, just run this script again and
# restart nginx (sudo systemctl reload nginx).
set -euo pipefail

IP="${1:-}"
OUT="${2:-$HOME/spark-certs}"
if [ -z "$IP" ]; then
  echo "Usage: bash deploy/make_cert.sh <PI-IP-ADDRESS> [output-folder]"
  echo "Example: bash deploy/make_cert.sh 192.168.1.50"
  exit 1
fi
if ! [[ "$IP" =~ ^[0-9]{1,3}(\.[0-9]{1,3}){3}$ ]]; then
  echo "'$IP' doesn't look like an IPv4 address (like 192.168.1.50)."
  exit 1
fi

mkdir -p "$OUT"
cd "$OUT"
umask 077

# 1) Certificate authority (created once; reused if it already exists)
if [ ! -f spark-ca.key ]; then
  openssl genrsa -out spark-ca.key 4096
  openssl req -x509 -new -nodes -key spark-ca.key -sha256 -days 1825 \
    -subj "/CN=SPARK Local CA" -out spark-ca.crt
fi

# 2) Server certificate signed by that authority
openssl genrsa -out spark.key 2048
openssl req -new -key spark.key -subj "/CN=spark.local" -out spark.csr
cat > spark-ext.cnf <<EOF
subjectAltName = DNS:spark.local, DNS:spark, IP:${IP}
basicConstraints = CA:FALSE
keyUsage = digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth
EOF
openssl x509 -req -in spark.csr -CA spark-ca.crt -CAkey spark-ca.key \
  -CAcreateserial -out spark.crt -days 397 -sha256 -extfile spark-ext.cnf

rm -f spark.csr spark-ext.cnf
chmod 644 spark-ca.crt spark.crt
chmod 600 spark.key spark-ca.key

echo
echo "Done. Files are in: $OUT"
ls -l "$OUT"
echo
echo "Next: install deploy/nginx_spark.conf (see deploy/SETUP_HTTPS.md)."
