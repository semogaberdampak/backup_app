-- ============================================================
-- TAKOM KASIR - Hardening Supabase (RLS + RPC verify PIN)
-- Jalankan di Supabase SQL Editor (schema public).
-- ============================================================

-- 1) Pastikan tabel users ada (username, role, pin_hash, salt)
--    (sesuaikan nama kolom dengan skema yang dipakai aplikasi)
-- create table if not exists public.users (
--   id bigserial primary key,
--   username text unique not null,
--   role text not null check (role in ('owner','admin','kasir')),
--   pin_hash text not null,
--   created_at timestamptz default now()
-- );

-- 2) Enable RLS pada tabel sensitif
alter table public.users enable row level security;
alter table public.master_produk enable row level security;

-- 3) RLS: anon boleh SELECT master_produk (untuk sync) tapi tidak menulis
drop policy if exists master_produk_read on public.master_produk;
create policy master_produk_read on public.master_produk
  for select to anon, authenticated
  using (true);

-- 4) RLS: users tidak bisa dibaca langsung oleh anon
drop policy if exists users_no_direct_read on public.users;
create policy users_no_direct_read on public.users
  for select to anon, authenticated
  using (false);

-- 5) RPC SECURITY DEFINER untuk verifikasi PIN (dipanggil dari client via anon key)
create or replace function public.verify_pin(p_username text, p_pin text)
returns json language plpgsql security definer as $$
declare
  v_user public.users%rowtype;
  v_hash text;
begin
  select * into v_user from public.users where username = p_username;
  if not found then
    return json_build_object('ok', false, 'error', 'user not found');
  end if;
  -- bandingkan hash (gunakan crypt() bila pin disimpan dengan pgcrypto)
  if v_user.pin_hash = crypt(p_pin, v_user.pin_hash) then
    return json_build_object('ok', true, 'username', v_user.username, 'role', v_user.role);
  end if;
  return json_build_object('ok', false, 'error', 'invalid pin');
end;
$$;

-- berikan izin eksekusi RPC ke anon
grant execute on function public.verify_pin(text, text) to anon, authenticated;

-- 6) (Opsional) fungsi sync master produk via RPC (lebih aman daripada direct SELECT)
create or replace function public.get_master_produk()
returns setof public.master_produk language sql security definer as $$
  select * from public.master_produk;
$$;
grant execute on function public.get_master_produk() to anon, authenticated;
