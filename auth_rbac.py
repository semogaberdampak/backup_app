"""RBAC: login PIN, session di memori, matriks peran + role dinamis (app_roles)."""
import hashlib
import logging
import os
import re
from dataclasses import dataclass
from typing import Dict, FrozenSet, Iterable, Optional

# Role sistem bawaan — role lain bisa dibuat owner lewat tabel app_roles
ROLES = ("kasir", "admin", "owner")

# kasir: transaksi + reprint; admin: + master/harga/laporan/setoran;
# owner: + kelola user/audit (owner selalu akses penuh, lihat sanitize_permissions)
PERMISSIONS = {
    "kasir": frozenset({"transaksi", "reprint"}),
    "admin": frozenset({"transaksi", "reprint", "master", "harga", "laporan", "setoran"}),
    "owner": frozenset(
        {"transaksi", "reprint", "master", "harga", "laporan", "setoran", "users", "audit"}
    ),
}

# (kunci, label) untuk checklist kelola role di UI owner — urutan tampilan
ALL_PERMISSIONS = (
    ("transaksi", "Transaksi kasir"),
    ("reprint", "Cetak ulang nota"),
    ("master", "Data master produk"),
    ("harga", "Ubah harga jual"),
    ("laporan", "Laporan & sheet"),
    ("setoran", "Kas & setoran"),
    ("users", "Kelola user & role"),
    ("audit", "Lihat audit log"),
)
PERMISSION_KEYS = frozenset(k for k, _ in ALL_PERMISSIONS)
ROLE_NAME_RE = re.compile(r"^[a-z0-9_]{2,20}$")


def hash_pin(pin: str, salt: str) -> str:
    """Hash PIN dengan salt tetap; kompatibel dengan data lama."""
    return hashlib.sha256(f"{salt}::{pin}".encode("utf-8")).hexdigest()


def pin_salt() -> str:
    return os.getenv("PIN_SALT", "takomers-default-salt-ganti-di-env")


@dataclass
class SessionUser:
    username: str
    role: str
    # Hak akses dinamis dari app_roles; None = pakai matriks statis PERMISSIONS
    permissions: Optional[FrozenSet[str]] = None

    def effective_permissions(self) -> FrozenSet[str]:
        if self.permissions is not None:
            return self.permissions
        return PERMISSIONS.get(self.role, frozenset())

    def can(self, action: str) -> bool:
        return action in self.effective_permissions()


def verify_pin(stored_hash: str, pin: str, salt: Optional[str] = None) -> bool:
    return stored_hash == hash_pin(pin, salt or pin_salt())


def default_owner_pin_hash() -> str:
    # PIN awal owner: 123456 (wajib diganti setelah login pertama)
    return hash_pin("123456", pin_salt())


def valid_role_name(name: str) -> bool:
    """Nama role: huruf kecil/angka/underscore, 2-20 karakter."""
    return bool(ROLE_NAME_RE.match(str(name or "")))


def sanitize_permissions(role: str, perms: Iterable) -> FrozenSet[str]:
    """Buang izin tak dikenal; role owner selalu akses penuh (anti lockout)."""
    if str(role) == "owner":
        return frozenset(PERMISSION_KEYS)
    return frozenset(str(p) for p in perms if str(p) in PERMISSION_KEYS)


def default_roles() -> Dict[str, FrozenSet[str]]:
    """Matriks role bawaan aplikasi (fallback saat offline)."""
    return {r: PERMISSIONS[r] for r in ROLES}


def load_roles(client=None) -> Dict[str, FrozenSet[str]]:
    """Muat role dinamis dari tabel app_roles; fallback default bila offline/gagal."""
    roles = default_roles()
    if client is None:
        return roles
    try:
        rows = client.table("app_roles").select("role,permissions").execute().data or []
        for row in rows:
            name = str(row.get("role") or "").strip()
            if name:
                roles[name] = sanitize_permissions(name, row.get("permissions") or [])
    except Exception as e:
        logging.error(f"Muat role dinamis gagal (pakai default): {e}", exc_info=True)
    return roles

