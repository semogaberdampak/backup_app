import pytest
from auth_rbac import SessionUser, hash_pin, verify_pin
from keuangan import hitung_setoran, ringkas_transaksi


class TestRBAC:
    def test_kasir_minimal(self):
        u = SessionUser(username="k", role="kasir")
        assert u.can("transaksi") and u.can("reprint")
        assert not u.can("harga") and not u.can("users")

    def test_admin_mid(self):
        u = SessionUser(username="a", role="admin")
        assert u.can("harga") and u.can("setoran")
        assert not u.can("users")

    def test_owner_full(self):
        u = SessionUser(username="o", role="owner")
        assert u.can("users") and u.can("audit")

    def test_pin_hash(self):
        h = hash_pin("123456", "salt-x")
        assert verify_pin(h, "123456", "salt-x")
        assert not verify_pin(h, "654321", "salt-x")


class TestKeuangan:
    def test_ringkas(self):
        r = ringkas_transaksi([{"total": 10000, "diskon": 500}, {"total": 5000}])
        assert r == {"total": 15000, "diskon": 500, "bersih": 14500}

    def test_setoran(self):
        s = hitung_setoran(100000, 50000, kas_awal=20000)
        assert s["omzet"] == 150000 and s["setor_tunai"] == 80000
