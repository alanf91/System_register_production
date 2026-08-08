# Alimentador PPCP por PDFs

Este pacote contém uma solução em Python para ler PDFs do relatório **Relatório de Atingimento por Setor/Peças** e alimentar a aba **5_ACOMPANHAMENTO** da planilha mãe.

## O que ele preenche

Na aba `5_ACOMPANHAMENTO`, o programa alimenta:

- **B - OP/Lote**: valor do campo `Lotes:` do PDF
- **C - Cliente/Pedido**: produto do cabeçalho do PDF, por exemplo `MAD 0707 (VERMELHO)`; é opcional na interface
- **D - Produto/Equipamento**: setor/equipamento do PDF, por exemplo `46 SECCIONADORA 1`
- **E - Codigo Peca**
- **F - Descricao Peca**
- **G - Operacao**: igual ao Produto/Equipamento
- **H - Setor**: igual ao Produto/Equipamento
- **I - Maquina/Posto**: igual ao Produto/Equipamento
- **J - Qtde Programada**: valor da coluna `Qtde.Peças` do PDF

A coluna **A - Data Programada** pode ser informada na interface. Se ficar vazia, o programa deixa em branco.

## Como instalar

1. Instale o Python 3.10 ou superior.
2. Abra o Prompt de Comando na pasta deste arquivo.
3. Rode:

```bash
pip install -r requirements_ppcp.txt
```

## Como usar com interface

```bash
python app_alimentar_ppcp.py
```

Depois:

1. Selecione a planilha mãe `PPCP_ACOMPAN.xlsx`.
2. Selecione um ou vários PDFs.
3. Informe a Data Programada, se desejar.
4. Clique em **PROCESSAR**.
5. O programa gera uma cópia da planilha com final `_atualizada_...xlsx`.

## Como usar sem interface

```bash
python app_alimentar_ppcp.py --sem-gui --planilha PPCP_ACOMPAN.xlsx --pdfs "SEC 1 - 10180,10148.pdf" --saida PPCP_ACOMPAN_atualizada.xlsx --data 30/06/2026
```

## Observações importantes

- O arquivo original não é sobrescrito automaticamente; o programa salva uma cópia atualizada.
- O programa copia formatação e fórmulas de uma linha modelo da própria aba, traduzindo as referências para as novas linhas.
- A opção **Evitar duplicados** usa a chave: lote + produto/equipamento + código da peça + descrição.
- Recomenda-se manter uma cópia de segurança da planilha mãe.
