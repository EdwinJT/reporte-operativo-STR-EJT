-- =============================================================================
-- SCRIPT DE CONFIGURACIÓN SUPABASE — CONTROL HORÓMETROS Y PRODUCCIÓN STRACON
-- Instrucciones: Copia y pega TODO este script en tu panel de Supabase:
-- "SQL Editor" -> "New query" -> Pegar -> Clic en "Run".
-- =============================================================================

-- 1. TABLA PRINCIPAL: public.registros_produccion
CREATE TABLE IF NOT EXISTS public.registros_produccion (
  id uuid PRIMARY KEY,
  codigo_equipo text NOT NULL,
  fecha date NOT NULL,
  turno text NOT NULL CHECK (turno IN ('DIA', 'NOCHE')),
  horometro_inicial numeric(10,1) NOT NULL,
  horometro_final numeric(10,1) NOT NULL,
  horas_trabajadas numeric(8,2) NOT NULL DEFAULT 0,
  material_fracturado numeric(10,2) NOT NULL DEFAULT 0,
  rendimiento numeric(8,2) NOT NULL DEFAULT 0,
  area_trabajo text DEFAULT '',
  operador text DEFAULT '',
  observaciones text DEFAULT '',
  hora_peru text DEFAULT '',
  creado_en timestamptz NOT NULL DEFAULT now(),
  actualizado_en timestamptz NOT NULL DEFAULT now(),
  dispositivo text DEFAULT '',
  eliminado boolean NOT NULL DEFAULT false,
  CONSTRAINT uq_registros_equipo_fecha_turno UNIQUE (codigo_equipo, fecha, turno)
);

-- Índices para consultas rápidas por equipo, fecha y turno
CREATE INDEX IF NOT EXISTS idx_registros_prod_equipo_fecha
  ON public.registros_produccion (codigo_equipo, fecha DESC);
CREATE INDEX IF NOT EXISTS idx_registros_prod_fecha
  ON public.registros_produccion (fecha DESC);

-- 2. CATÁLOGO COMPARTIDO DE EQUIPOS
CREATE TABLE IF NOT EXISTS public.equipos_produccion (
  codigo text PRIMARY KEY,
  categoria text DEFAULT '',
  actualizado_en timestamptz NOT NULL DEFAULT now()
);

-- 3. CATÁLOGO COMPARTIDO DE ÁREAS DE TRABAJO
CREATE TABLE IF NOT EXISTS public.areas_trabajo (
  nombre text PRIMARY KEY,
  actualizado_en timestamptz NOT NULL DEFAULT now()
);

-- 4. ACTIVAR ROW LEVEL SECURITY (RLS) Y POLÍTICAS MULTI-DISPOSITIVO EN CAMPO
ALTER TABLE public.registros_produccion ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.equipos_produccion ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.areas_trabajo ENABLE ROW LEVEL SECURITY;

-- Limpiar políticas previas en registros_produccion
DROP POLICY IF EXISTS "registros_prod_select_all" ON public.registros_produccion;
DROP POLICY IF EXISTS "registros_prod_insert_all" ON public.registros_produccion;
DROP POLICY IF EXISTS "registros_prod_update_all" ON public.registros_produccion;
DROP POLICY IF EXISTS "registros_prod_delete_all" ON public.registros_produccion;

CREATE POLICY "registros_prod_select_all"
  ON public.registros_produccion FOR SELECT TO anon, authenticated USING (true);
CREATE POLICY "registros_prod_insert_all"
  ON public.registros_produccion FOR INSERT TO anon, authenticated WITH CHECK (true);
CREATE POLICY "registros_prod_update_all"
  ON public.registros_produccion FOR UPDATE TO anon, authenticated USING (true) WITH CHECK (true);
CREATE POLICY "registros_prod_delete_all"
  ON public.registros_produccion FOR DELETE TO anon, authenticated USING (true);

-- Limpiar políticas previas en equipos_produccion
DROP POLICY IF EXISTS "equipos_prod_select_all" ON public.equipos_produccion;
DROP POLICY IF EXISTS "equipos_prod_insert_all" ON public.equipos_produccion;
DROP POLICY IF EXISTS "equipos_prod_update_all" ON public.equipos_produccion;
DROP POLICY IF EXISTS "equipos_prod_delete_all" ON public.equipos_produccion;

CREATE POLICY "equipos_prod_select_all"
  ON public.equipos_produccion FOR SELECT TO anon, authenticated USING (true);
CREATE POLICY "equipos_prod_insert_all"
  ON public.equipos_produccion FOR INSERT TO anon, authenticated WITH CHECK (true);
CREATE POLICY "equipos_prod_update_all"
  ON public.equipos_produccion FOR UPDATE TO anon, authenticated USING (true) WITH CHECK (true);
CREATE POLICY "equipos_prod_delete_all"
  ON public.equipos_produccion FOR DELETE TO anon, authenticated USING (true);

-- Limpiar políticas previas en areas_trabajo
DROP POLICY IF EXISTS "areas_trabajo_select_all" ON public.areas_trabajo;
DROP POLICY IF EXISTS "areas_trabajo_insert_all" ON public.areas_trabajo;
DROP POLICY IF EXISTS "areas_trabajo_update_all" ON public.areas_trabajo;
DROP POLICY IF EXISTS "areas_trabajo_delete_all" ON public.areas_trabajo;

CREATE POLICY "areas_trabajo_select_all"
  ON public.areas_trabajo FOR SELECT TO anon, authenticated USING (true);
CREATE POLICY "areas_trabajo_insert_all"
  ON public.areas_trabajo FOR INSERT TO anon, authenticated WITH CHECK (true);
CREATE POLICY "areas_trabajo_update_all"
  ON public.areas_trabajo FOR UPDATE TO anon, authenticated USING (true) WITH CHECK (true);
CREATE POLICY "areas_trabajo_delete_all"
  ON public.areas_trabajo FOR DELETE TO anon, authenticated USING (true);
