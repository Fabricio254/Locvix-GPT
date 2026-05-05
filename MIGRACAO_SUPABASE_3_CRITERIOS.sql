-- ═════════════════════════════════════════════════════════════════════════════════
-- MIGRAÇÃO: Adição de 3 Critérios de Manutenção (Horímetro + Hodômetro + Período)
-- ═════════════════════════════════════════════════════════════════════════════════
-- 
-- ANTES DE EXECUTAR:
-- 1. Abra o Supabase Console (https://app.supabase.com)
-- 2. Acesse seu projeto
-- 3. Vá para "SQL Editor"
-- 4. Cole TODO O CONTEÚDO ABAIXO
-- 5. Clique em "Run"
-- 
-- Esta migração adiciona 3 novos campos à tabela `manutencoes_equipamentos`:
-- - hodometro_ultima_manutencao (NUMERIC): Hodômetro (km) no momento da última manutenção
-- - intervalo_km (NUMERIC): Intervalo de manutenção em quilômetros
-- - periodo_dias (INTEGER): Intervalo de manutenção em dias (ex: 90 dias)
-- 
-- ═════════════════════════════════════════════════════════════════════════════════

-- Verificar estrutura atual da tabela
-- SELECT column_name, data_type, is_nullable 
-- FROM information_schema.columns 
-- WHERE table_name = 'manutencoes_equipamentos';

-- Adicionar coluna: hodometro_ultima_manutencao
ALTER TABLE public.manutencoes_equipamentos
ADD COLUMN IF NOT EXISTS hodometro_ultima_manutencao NUMERIC(12, 2);

COMMENT ON COLUMN public.manutencoes_equipamentos.hodometro_ultima_manutencao IS 'Hodômetro (km) no momento da última manutenção';

-- Adicionar coluna: intervalo_km
ALTER TABLE public.manutencoes_equipamentos
ADD COLUMN IF NOT EXISTS intervalo_km NUMERIC(12, 2);

COMMENT ON COLUMN public.manutencoes_equipamentos.intervalo_km IS 'Intervalo de manutenção em quilômetros (km)';

-- Adicionar coluna: periodo_dias
ALTER TABLE public.manutencoes_equipamentos
ADD COLUMN IF NOT EXISTS periodo_dias INTEGER;

COMMENT ON COLUMN public.manutencoes_equipamentos.periodo_dias IS 'Intervalo de manutenção em dias (ex: 90 dias)';

-- Adicionar coluna: ultima_manutencao (data)
ALTER TABLE public.manutencoes_equipamentos
ADD COLUMN IF NOT EXISTS ultima_manutencao DATE;

COMMENT ON COLUMN public.manutencoes_equipamentos.ultima_manutencao IS 'Data da última manutenção realizada';

-- ═════════════════════════════════════════════════════════════════════════════════
-- RESULTADO ESPERADO:
-- 4 novas colunas adicionadas à tabela (ou nenhuma alteração se já existem)
-- ═════════════════════════════════════════════════════════════════════════════════
