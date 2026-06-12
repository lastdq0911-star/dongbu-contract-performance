-- ════════════════════════════════════════════════════════════
-- 컨테이너운송 전용 컬럼 추가
-- Supabase 대시보드 → SQL Editor 에 붙여넣고 [Run] 한 번만 실행하면 됩니다.
-- (이미 있으면 IF NOT EXISTS 로 무시되므로 여러 번 실행해도 안전)
--
-- [실행 방법]
-- 1) https://supabase.com/dashboard 접속 → 해당 프로젝트 선택
-- 2) 좌측 메뉴 "SQL Editor" 클릭 → "New query"
-- 3) 이 파일 내용 전체를 복사해서 붙여넣기
-- 4) 우측 상단 "Run" (또는 Ctrl/Cmd+Enter) 클릭
-- 5) "Success. No rows returned" 메시지가 보이면 완료
-- ════════════════════════════════════════════════════════════

ALTER TABLE performance_records
  ADD COLUMN IF NOT EXISTS extra_partner   text,    -- 추가협력사
  ADD COLUMN IF NOT EXISTS work_client     text,    -- 작업화주명
  ADD COLUMN IF NOT EXISTS service_detail  text,    -- 세부용역(용역)
  ADD COLUMN IF NOT EXISTS surcharge_name  text,    -- 할증명
  ADD COLUMN IF NOT EXISTS sz_code         text,    -- SZ
  ADD COLUMN IF NOT EXISTS billing_amount  numeric, -- 청구금액
  ADD COLUMN IF NOT EXISTS billing_client  text,    -- 청구화주
  ADD COLUMN IF NOT EXISTS drop_point      text;    -- 하차지
