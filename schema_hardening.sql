-- Execute uma vez no MySQL de produção depois de encerrar entradas duplicadas existentes.
-- NULLs não colidem em índice UNIQUE; assim, visitas concluídas continuam no histórico.
ALTER TABLE entrada
    ADD COLUMN visitante_em_visita BIGINT
        GENERATED ALWAYS AS (CASE WHEN data_hora_saida IS NULL THEN visitante_id ELSE NULL END) STORED,
    ADD UNIQUE INDEX ux_entrada_visitante_em_visita (visitante_em_visita);
