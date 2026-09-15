#!/bin/bash
# HP Victus Hall Sensor Fix for Linux
# Disables false lid-closed detection

set -e

echo "=== HP Victus Hall Sensor Fix ==="

# Check if running as root or with sudo
if [ "$EUID" -ne 0 ]; then
    echo "This script needs root. Run with: sudo bash $0"
    exit 1
fi

echo "[1/3] Applying fix immediately..."
echo "open" > /sys/module/button/parameters/lid_init_state
echo "  Done. Current value: $(cat /sys/module/button/parameters/lid_init_state)"

echo "[2/3] Making permanent across reboots..."

# systemd-boot (CachyOS)
if ls /boot/loader/entries/*.conf &>/dev/null; then
    for entry in /boot/loader/entries/*.conf; do
        if grep -q "button.lid_init_state=open" "$entry"; then
            echo "  $entry already configured."
        else
            sed -i 's/^options /options /' "$entry"
            # Add param to options line if not present
            sed -i '/^options / s/$/ button.lid_init_state=open/' "$entry"
            echo "  Updated: $entry"
        fi
    done
    echo "  Boot entries updated."
else
    echo "  WARNING: Could not find systemd-boot entries."
    echo "  Add 'button.lid_init_state=open' to your boot options manually."
fi

echo "[3/3] Verifying..."
echo "  Lid init state: $(cat /sys/module/button/parameters/lid_init_state)"
echo ""
echo "=== Done! Reboot to fully apply. ==="
echo "After reboot, verify with:"
echo "  cat /sys/module/button/parameters/lid_init_state"
