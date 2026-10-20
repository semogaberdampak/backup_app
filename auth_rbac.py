"""RBAC: login PIN, session di memori, matriks peran."""
import hashlib
import os
from dataclasses import dataclass
from typing import Optional

ROLES = ("kasir", "admin", "owner")

# kasir: transaksi + reprint; admin: + master/harga/laporan/setoran;
# owner: + kelola user/audit
PERMISSIONS = {
    "kasir": frozenset({"transaksi", "reprint"}),
    "admin": frozenset({"transaksi", "reprint", "master", "harga", "laporan", "setoran"}),
    "owner": frozenset(
        {"transaksi", "reprint", "master", "harga", "laporan", "setoran", "users", "audit"}
    ),
}


def hash_pin(pin: str, salt: str) -> str:
    return hashlib.sha256(f"{salt}::{pin}".encode("utf-8")).hexdigest()


def pin_salt() -> str:
    return os.getenv("PIN_SALT", "takomers-default-salt-ganti-di-env")


@dataclass
class SessionUser:
    username: str
    role: str

    def can(self, action: str) -> bool:
        return action in PERMISSIONS.get(self.role, frozenset())


def verify_pin(stored_hash: str, pin: str, salt: Optional[str] = None) -> bool:
    return stored_hash == hash_pin(pin, salt or pin_salt())


def default_owner_pin_hash() -> str:
    # PIN awal owner: 123456 (wajib diganti setelah login pertama)
    return hash_pin("123456", pin_salt())
