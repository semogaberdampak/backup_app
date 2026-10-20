import pytest
from auth_rbac import (
    ALL_PERMISSIONS,
    PERMISSIONS,
    PERMISSION_KEYS,
    SessionUser,
    default_roles,
    hash_pin,
    load_roles,
    sanitize_permissions,
    valid_role_name,
    verify_pin,
)
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


class TestRoleDinamis:
    def test_kunci_izin_unik_dan_sesuai_matriks(self):
        kunci = [k for k, _ in ALL_PERMISSIONS]
        assert len(kunci) == len(set(kunci))
        assert set(kunci) == PERMISSION_KEYS == set().union(*PERMISSIONS.values())

    def test_session_custom_permissions(self):
        u = SessionUser(username="budi", role="kurir",
                        permissions=frozenset({"laporan"}))
        assert u.can("laporan")
        assert not u.can("transaksi") and not u.can("users")

    def test_session_fallback_statis(self):
        u = SessionUser(username="k", role="kasir")
        assert u.can("transaksi") and not u.can("harga")

    def test_sanitize_buang_izin_tak_dikenal(self):
        hasil = sanitize_permissions("supervisor", ["laporan", "hantu", 123])
        assert hasil == frozenset({"laporan"})

    def test_owner_selalu_penuh(self):
        assert sanitize_permissions("owner", ["transaksi"]) == PERMISSION_KEYS

    def test_valid_role_name(self):
        assert valid_role_name("supervisor") and valid_role_name("kurir_2")
        assert not valid_role_name("Admin")
        assert not valid_role_name("a")
        assert not valid_role_name("ada spasi")
        assert not valid_role_name("x" * 21)

    def test_default_roles(self):
        r = default_roles()
        assert set(r) == {"kasir", "admin", "owner"}
        assert r["owner"] == PERMISSION_KEYS

    def test_load_roles_offline_fallback(self):
        assert load_roles(None) == default_roles()

    def test_load_roles_dari_client(self):
        class _Res:
            data = [{"role": "kurir", "permissions": ["transaksi", "hantu"]}]

        class _T:
            def select(self, *_a, **_k):
                return self

            def execute(self):
                return _Res()

        class _Client:
            def table(self, _nama):
                return _T()

        r = load_roles(_Client())
        assert r["kurir"] == frozenset({"transaksi"})
        assert "kasir" in r and r["owner"] == PERMISSION_KEYS

    def test_load_roles_client_gagal_fallback(self):
        class _Client:
            def table(self, _nama):
                raise RuntimeError("db down")

        assert load_roles(_Client()) == default_roles()
