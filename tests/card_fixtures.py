def card_tables(nome="Ana Souza"):
    return [
        [
            ["NOME DA EMPRESA: Empresa Teste", None],
            ["CNPJ DA EMPRESA: 12345678000190", None],
            [f"NOME DO FUNCION\ufffdRIO: {nome}", "CPF DO FUNCION\ufffdRIO: 12345678901"],
            [
                "PIS DO FUNCION\ufffdRIO: 12345678900\nNOME DO CARGO: Operadora",
                "DATA DE ADMISS\ufffdO DO FUNCION\ufffdRIO: 08/01/2026\nN\ufffdMERO DE MATR\ufffdCULA: 7",
            ],
        ],
        [
            ["DIA", "PREVISTO", "ENT. 1", "SAI. 1", "ENT. 2", "SAI. 2", "TOTAL NORMAIS", "TOTAL NOTURNO", "DIA FALTA", "FALTA E ATRASO", "ABONO"],
            ["01/08/2026 - SAB", "", "Folga", None, None, None, "", "", "", "", ""],
            ["02/08/2026 - DOM", "00:00-02:00 03:00-08:00", "23:48 (M)", "02:00 (P)", "03:00 (P)", "08:07 (M)", "07:34", "05:57", "", "", ""],
            ["03/08/2026 - SEG", "08:00-12:00 13:00-16:00", "Falta", "Falta", "Falta", "Falta", "", "", "1", "", ""],
            ["04/08/2026 - TER", "08:00-12:00 13:00-16:00", "Abonar aus\ufffdncia por licen\ufffda", "", "", "", "", "", "", "", "07:00"],
            ["05/08/2026 - QUA", "08:00-12:00 13:00-16:00", "Atestado M\ufffddico", "", "", "", "", "", "", "", ""],
            ["TOTAIS", "", "", "", "", "", "07:34", "05:57", "1", "", ""],
        ],
    ]
