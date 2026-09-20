create table if not exists public.calendar_events (
  date text primary key,
  title text not null,
  class_name text not null default '',
  deleted boolean not null default false,
  updated_at timestamptz not null default now()
);

alter table public.calendar_events enable row level security;

-- The Vercel API uses the service-role key, so browser clients never access Supabase directly.
