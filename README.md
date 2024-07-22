# Script de Web Scraping e Automação
Este script em Python, desenvolvido para uso exclusivo da empresa Eizeitz, automatiza o processo de web scraping, compilação de dados em planilhas Excel e distribuição por e-mail. O script é executado uma vez ao dia em um horário definido pelo usuário, extraindo dados de um site especificado, organizando-os em duas planilhas Excel e enviando os resultados para um destinatário designado.

## Funcionalidades
Web Scraping: Extrai dados específicos de um site utilizando as bibliotecas BeautifulSoup e Requests.
Compilação em Excel: Organiza os dados extraídos em duas planilhas Excel separadas utilizando as bibliotecas Pandas e OpenPyXL.
E-mail Automatizado: Envia o arquivo Excel compilado por e-mail utilizando as bibliotecas smtplib e email.
Execução Agendada: Executa automaticamente em um horário especificado pelo usuário todos os dias.
