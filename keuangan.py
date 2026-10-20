"""Keuangan: agregat transaksi + setoran + arus kas (pure + wrapper Supabase)."""
import logging
from datetime import date
from typing import Any, Dict, List, Optional


def ringkas_transaksi(items: List[Dict[str, Any]]) -> Dict[str, float]:
    total = sum(float(i.get("total", 0) or 0) for i in items)
    diskon = sum(float(i.get("diskon", 0) or 0) for i in items)
    return {"total": total, "diskon": diskon, "bersih": total - diskon}


def hitung_setoran(
    total_tunai: float, total_nontunai: float, kas_awal: float = 0.0
) -> Dict[str, float]:
    omzet = total_tunai + total_nontunai
    return {
        "omzet": omzet,
        "total_tunai": total_tunai,
        "total_nontunai": total_nontunai,
        "setor_tunai": max(total_tunai - kas_awal, 0.0),
    }


def catat_setoran(supabase_client, tanggal: str, tunai: float, nontunai: float,
                  disetor_oleh: str = "") -> Optional[Dict[str, Any]]:
    """Insert ke kas_setoran. Return row / None bila offline."""
    if supabase_client is None:
        return None
    try:
        res = (
            supabase_client.table("kas_setoran")
            .upsert(
                {
                    "tanggal": tanggal,
                    "total_tunai": float(tunai),
                    "total_nontunai": float(nontunai),
                    "disetor_oleh": disetor_oleh,
                },
                on_conflict="tanggal",
            )
            .execute()
        )
        return (res.data or [None])[0]
    except Exception as e:
        logging.error(f"Gagal catat setoran: {e}", exc_info=True)
        return None


def catat_arus_kas(supabase_client, jenis: str, kategori: str, nominal: float,
                   keterangan: str = "", dibuat_oleh: str = "",
                   tanggal: Optional[str] = None) -> Optional[Dict[str, Any]]:
    if supabase_client is None or jenis not in ("masuk", "keluar"):
        return None
    try:
        payload: Dict[str, Any] = {
            "jenis": jenis,
            "kategori": kategori,
            "nominal": float(nominal),
            "keterangan": keterangan,
            "dibuat_oleh": dibuat_oleh,
        }
        if tanggal:
            payload["tanggal"] = tanggal
        else:
            payload["tanggal"] = date.today().isoformat()
        res = supabase_client.table("arus_kas").insert(payload).execute()
        return (res.data or [None])[0]
    except Exception as e:
        logging.error(f"Gagal catat arus kas: {e}", exc_info=True)
        return None
