# Baixador de XML de NF-e

Aplicação web **local** (roda no seu computador, usada pelo navegador) que baixa os XMLs de NF-e
do site consultadanfe.com a partir das chaves de acesso de uma planilha.

## Requisitos
- **Python 3.9+** (https://www.python.org/downloads/ — no Windows marque "Add python.exe to PATH")
- **Google Chrome** instalado e internet (o driver do Chrome é baixado automaticamente na 1ª vez)

## Como usar
1. **Windows:** dê dois cliques em `iniciar.bat` (ou, no Anaconda Prompt, `python app.py`)  
   **Mac/Linux:** no terminal, `./iniciar.sh`
   (precisa de flask, pandas, openpyxl e selenium; se faltar algum, o programa avisa qual instalar)
2. O navegador abre em `http://127.0.0.1:8765`. Então:
   1. Escolha a planilha (.xlsx/.csv) e confirme a coluna da chave (ex.: `CHAVE_ACESSO_NOTA`)
   2. Informe a linha inicial e final (ex.: 1753 a 1802). A linha 1 é a primeira nota (o cabeçalho não conta)
   3. Escolha a pasta onde salvar os XMLs
   4. Clique em **Executar**
3. Uma janela do Chrome abre sozinha. **Não a feche.** Se aparecer a verificação anti-robô (Cloudflare),
   resolva na própria janela e o programa continua.

## Etapa 2 — Importar no SISGI (aba "Importar no SISGI")
1. Escolha a pasta com os XMLs (o botão "Usar a pasta do download" já preenche).
2. Opcional: filtre por mês de competência ou limite a quantidade (bom para testar).
3. Confira o endereço do SISGI (o padrão é o de **homologação**, `sisgi.hmg…`) e clique em **Importar**.
4. O Chrome abre no SISGI: **faça o login você mesmo**. O programa continua sozinho quando a lista de despesas aparecer.
5. Os XMLs são agrupados pelo mês de emissão e enviados em lotes de até 20. Ao final, clique em OK no aviso do Chrome.
- O que já foi enviado fica em `sisgi_importados.txt` dentro da pasta dos XMLs e não é reenviado
  (marque "Reenviar também os já enviados" para ignorar esse histórico).
- *Apenas simular* mostra quantos arquivos há por mês, sem abrir o Chrome.

## Dicas
- Rode em lotes pequenos (30–50 notas) para evitar bloqueio. As pausas podem ser ajustadas em *Opções avançadas*.
- Chaves já baixadas (`<chave>.xml` na pasta) são puladas — para completar falhas, execute de novo o mesmo intervalo.
- *Apenas simular* mostra o que seria feito sem abrir o Chrome nem baixar nada.
- Para usar em outro computador, copie esta pasta inteira e rode o iniciador de novo.
- Só este computador acessa a aplicação (escuta em 127.0.0.1). Para encerrar: botão "Encerrar aplicação" ou feche a janela preta.
