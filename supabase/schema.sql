-- Run this once in the Supabase SQL editor (Project > SQL Editor > New query).

create extension if not exists pgcrypto;

create table if not exists prospects (
    id uuid primary key default gen_random_uuid(),
    name text not null,
    board text not null check (board in ('OFFENSE', 'DEFENSE')),
    position text not null,
    school text default '',
    draft_class_year int not null,
    entry_year int,
    agents text default '',
    created_by uuid references auth.users(id),
    created_at timestamptz not null default now()
  );

create table if not exists grades (
    id uuid primary key default gen_random_uuid(),
    prospect_id uuid not null references prospects(id) on delete cascade,
    scout text not null,
    grade numeric not null,
    created_by uuid references auth.users(id),
    created_at timestamptz not null default now()
  );

create index if not exists prospects_board_year_idx on prospects (board, draft_class_year);
create index if not exists grades_prospect_idx on grades (prospect_id);

alter table prospects enable row level security;
alter table grades enable row level security;

-- Any signed-in team member (invited via Supabase Auth) can read and write.
-- There is no public sign-up flow in the app, so the only people who can
-- reach these tables are people you've explicitly invited.
create policy "Authenticated team members can read prospects"
  on prospects for select
  using (auth.role() = 'authenticated');

create policy "Authenticated team members can write prospects"
  on prospects for insert
  with check (auth.role() = 'authenticated');

create policy "Authenticated team members can update prospects"
  on prospects for update
  using (auth.role() = 'authenticated');

create policy "Authenticated team members can delete prospects"
  on prospects for delete
  using (auth.role() = 'authenticated');

create policy "Authenticated team members can read grades"
  on grades for select
  using (auth.role() = 'authenticated');

create policy "Authenticated team members can write grades"
  on grades for insert
  with check (auth.role() = 'authenticated');

create policy "Authenticated team members can delete grades"
  on grades for delete
  using (auth.role() = 'authenticated');

-- Enables live updates so teammates see each other's edits without refreshing.
alter publication supabase_realtime add table prospects;
alter publication supabase_realtime add table grades;
