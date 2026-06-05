-- ════════════════════════════════════════════════════════════
-- 성능 인덱스 (데이터가 많아질 때 조회 속도 향상)
-- Supabase 대시보드 → SQL Editor 에 붙여넣고 [Run] 한 번만 실행하면 됩니다.
-- (이미 있으면 IF NOT EXISTS 로 무시되므로 여러 번 실행해도 안전)
-- ════════════════════════════════════════════════════════════

-- 1) 연도/월/지사 복합 인덱스 — 대시보드·필터의 핵심 조회 경로
CREATE INDEX IF NOT EXISTS idx_records_year_month_branch
  ON performance_records (year, month, branch);

-- 2) 배치 단위 조회·삭제 가속 (배치관리 화면, 배치 삭제)
CREATE INDEX IF NOT EXISTS idx_records_batch_id
  ON performance_records (batch_id);

-- 3) 협력사별 화면 가속
CREATE INDEX IF NOT EXISTS idx_records_partner
  ON performance_records (partner_name);

-- 4) 고객별 화면 가속
CREATE INDEX IF NOT EXISTS idx_records_client
  ON performance_records (client);

-- 5) 연도 우선 로딩(최근 연도 먼저 불러오기) 가속
CREATE INDEX IF NOT EXISTS idx_records_year
  ON performance_records (year);

-- batches 테이블도 연/월/지사로 자주 조회됨
CREATE INDEX IF NOT EXISTS idx_batches_year_month_branch
  ON batches (year, month, branch);
