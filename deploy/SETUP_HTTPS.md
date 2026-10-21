# SPARK on the Raspberry Pi: `spark.local` + HTTPS

Goal: students open **https://spark.local** (no IP address to type), and
everything they send, including passwords, is encrypted so Wireshark only shows
scrambled data.

Do these steps **on the Raspberry Pi** (over SSH). Replace `spark` with your
username if it is different (and fix the paths in `nginx_spark.conf` and
`spark.service` to match).

---

## Step 1: Give the Pi a fixed IP (in the router)

Open your router's admin page, find **DHCP reservation / Static lease / Address
reservation**, and reserve one IP for the Pi (example: `192.168.1.50`).
The certificate includes this IP, so it must not change.

Check the Pi's current IP: `hostname -I`

## Step 2: Name the Pi `spark`

```bash
sudo hostnamectl set-hostname spark
sudo apt install avahi-daemon -y     # usually already installed
sudo reboot
```

After it reboots, from a Windows PC on the same router: `ping spark.local`
should reply with the Pi's IP.

Notes:
- Windows 10/11, iPhone/iPad and Mac open `spark.local` with no setup.
- Some Android phones can't open `.local` names. For those, use the IP
  (`https://192.168.1.50`) , it works too.
- Typing just `spark` in a browser starts a Google search, so tell students to
  type `spark.local` (or give them a QR code / printed handout of
  `https://spark.local`).
- If your router has a "local DNS / hostname" setting, add `spark` -> Pi IP.
  Then `https://spark` also works.

## Step 3: Pull the new code and install packages

```bash
cd ~/capss-main
git pull
source venv/bin/activate
pip install -r requirements.txt
```

## Step 4: Make the certificates

```bash
bash deploy/make_cert.sh 192.168.1.50      # <- the Pi's IP from Step 1
```

This creates the folder `~/spark-certs` containing `spark.crt` / `spark.key`
(for nginx) and `spark-ca.crt` (optional, see "No warning" below).

## Step 5: Turn on HTTPS in nginx

```bash
sudo cp deploy/nginx_spark.conf /etc/nginx/sites-available/spark
sudo ln -sf /etc/nginx/sites-available/spark /etc/nginx/sites-enabled/spark
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl reload nginx
```

`nginx -t` must say "syntax is ok" and "test is successful". If it complains
about a path, fix the `/home/spark/...` lines in the file.

## Step 6: Update the SPARK service

```bash
sudo cp deploy/spark.service /etc/systemd/system/spark.service
sudo systemctl daemon-reload
sudo systemctl restart spark
sudo systemctl status spark          # should say "active (running)"
```

This turns on two settings: `SPARK_HTTPS=1` (login cookie only travels over
HTTPS) and `SPARK_BEHIND_PROXY=1` (each student is seen with their own IP, not
the Pi's, so the login rate limit doesn't block a whole class at once).

## Step 7: Test

From another device: open `http://spark.local`. It should jump to
`https://spark.local`.

---

## The "Not secure / Your connection is not private" warning

The certificate is made by the Pi itself, so browsers don't recognize who
issued it. The connection **is still encrypted**. The warning only says
"I can't verify the issuer." On each device, once:

- Chrome / Edge: **Advanced** -> **Continue to spark.local (unsafe)**
- Safari (iPhone): **Show Details** -> **visit this website**
- Firefox: **Advanced** -> **Accept the Risk and Continue**

### Optional: no warning at all

Install the Pi's certificate authority on the device, one time:

1. On that device open `http://spark.local/spark-ca.crt` (it downloads).
2. Install it:
   - **Windows:** double-click -> Install Certificate -> Local Machine ->
     "Place all certificates in the following store" -> **Trusted Root
     Certification Authorities**.
   - **Android:** Settings -> Security -> Encryption & credentials -> Install a
     certificate -> **CA certificate** -> pick the file.
   - **iPhone/iPad:** open the file -> Settings -> "Profile Downloaded" ->
     Install, then Settings -> General -> About -> **Certificate Trust
     Settings** -> turn on **SPARK Local CA**.
3. Restart the browser.

Doing this for every student device is a lot of work, so for exam day the
one-time "Proceed" click is usually easier. Teacher/admin laptops can install
the CA so the main devices show a clean padlock.

(A completely warning-free setup for everyone needs a real domain name and a
Let's Encrypt certificate. It costs a little yearly and needs internet on the
Pi occasionally. Not required for the defense.)

---

## Proving it works (for the panel)

1. Open Wireshark on a PC on the same network, start capturing.
2. Log in to SPARK on another device.
3. Filter: `http.request.method == POST`: **nothing appears** (no readable
   login). Filter `tls` and you'll see only encrypted "Application Data".
4. For comparison, the old `http://<ip>:5000` (if still reachable) shows the
   username and password in plain text. You can demonstrate "before and after".

## Maintenance

- Certificate lasts 397 days. To renew: `bash deploy/make_cert.sh <PI-IP>` then
  `sudo systemctl reload nginx`. (Renewing keeps the same authority, so devices
  that installed it need nothing new.)
- If the Pi's IP changes, run `make_cert.sh` again with the new IP.
- **Never share or commit** `spark-ca.key` or `spark.key`. They live in
  `~/spark-certs`, outside the project folder, so `git add .` won't pick them up.
- Running locally on Windows (`python app.py`) is unchanged: HTTPS settings are
  only on when the systemd service sets those two environment variables.
