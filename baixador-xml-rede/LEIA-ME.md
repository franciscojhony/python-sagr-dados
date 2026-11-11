# Baixador de XML de NF-e — servidor + agentes

O SERVIDOR guarda as planilhas, distribui as chaves e recebe os XMLs. O AGENTE roda no computador
de cada usuário: é ele que abre o Chrome ali, baixa e envia o XML ao servidor.

## 1) No servidor (uma vez)
1. Iniciar: `iniciar-servidor.bat` (ou `python app.py`). Precisa de: flask, pandas, openpyxl. (Não precisa de Chrome.)
2. Ajuste, no topo do `app.py` (bloco CONFIGURAÇÕES): MAX_USUARIOS, PASTA_XML, ACESSO_PELA_REDE.
3. Gere o Python portátil do Agente (precisa de internet): `preparar-agente.bat`
   Isso cria a pasta `agente-portatil\python`. Sem ela, o Agente só roda em computadores que já tenham Python.
4. Libere o Python no firewall do Windows (Rede privada) na primeira execução.

## 2) Em cada computador de usuário (uma vez)
1. Abra `http://IP-DO-SERVIDOR:8765` e clique em "Baixar o Agente" (agente-baixador.zip, já com o endereço do servidor).
2. Extraia em qualquer pasta. Precisa do Google Chrome e de internet (o driver do Chrome é baixado na 1ª vez).

## 3) Uso diário
1. O usuário abre `iniciar-agente.bat` (deixa a janela aberta) e a página do servidor no navegador.
2. Na página: escolhe a planilha (do próprio computador), informa as linhas e clica em Executar.
3. O Chrome abre NO COMPUTADOR DO USUÁRIO. Se pedir a verificação anti-robô, ele resolve ali mesmo.
4. Os XMLs chegam na pasta do servidor; o andamento aparece na página.

## Como o servidor liga página e Agente
Pelo IP do computador: o Agente que está no mesmo computador da página assume a tarefa daquela pessoa.
Por isso, cada pessoa usa o seu próprio computador. Se a tarefa ficar 10 min sem Agente, ou o Agente
ficar 3 min sem dar sinal, ela é cancelada (TEMPO_ESPERA_AGENTE / TEMPO_SEM_SINAL no app.py).
