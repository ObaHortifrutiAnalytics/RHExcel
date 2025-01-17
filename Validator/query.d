SELECT 
C.NROEMPRESA CODLOJA, C.NROPDV CODCAIXA, C.DATA_VENDA DATVENDA, C.NUMERODF NUMTICKET, C.CODOPERADOR CODOPERADOR, C.CPFCNPJ IDCLIENTE, COUNT(C.SEQPRODUTO) QTDVENDA, SUM(C.Valor_Liquido) VALVENDA, C.IDCUPOM 
FROM DW.IN_VENDA_SEQCUPOM C 
WHERE 
TRUNC(C.DATA_VENDA) =   trunc(SYSDATE-1)
AND C.CATEGORIA_1 IN ('MERCEARIA', 'ACESSORIO', 'BISCOITERIA', 'PEIXARIA', 'AUTOSSERVICO', 'PADARIA', 'FLV', 'ADEGA', 'FLORICULTURA', 'FRIOS', 'ACOUGUE', 'ROTISSERIA') 
GROUP BY C.NROEMPRESA, C.NROPDV, C.DATA_VENDA, C.NUMERODF, C.CODOPERADOR, C.CPFCNPJ, C.IDCUPOM 
ORDER BY 1, 2, 3; 

