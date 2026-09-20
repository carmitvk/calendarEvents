create table if not exists public.calendar_events (
  date date primary key,
  title text not null,
  class_name text not null,
  created_at timestamptz not null default now()
);

alter table public.calendar_events enable row level security;
