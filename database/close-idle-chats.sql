-- Applied to the production Supabase project through migration
-- close_website_chats_after_five_idle_hours. Retains all chat history.
create extension if not exists pg_cron;
select cron.schedule(
  'close-idle-website-chats',
  '* * * * *',
  $job$
    update public.conversations as c
    set status = 'closed', updated_at = now()
    where c.status in ('open', 'pending')
      and coalesce(c.last_message_at, c.created_at) <= now() - interval '5 hours'
      and not exists (
        select 1 from public.messages m
        where m.conversation_id = c.id
          and m.created_at > now() - interval '5 hours'
      );
  $job$
);
