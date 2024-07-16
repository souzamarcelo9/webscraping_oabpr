import logging

from bot import Bot
from ext.functions import create_excel_file, verificar_ffmpeg

instalar_ffmpeg = input("Deseja instalar o FFMPEG antes? (sim/nao): ")

if 's' in instalar_ffmpeg.lower().strip():
    verificar_ffmpeg()

# Instancia o robô que executa a automação
bot = Bot(False)
bot.open_oab()

# Cria o arquivo Excel onde será salvo as informações
create_excel_file()

# Obtém o total de páginas da lista de advogados
total_paginas = bot.get_last_page()
total_advogados = total_paginas * 15

# Percorre todos os advogados de todas as páginas da lista de advogados
for pagina in range(1, total_paginas+1):
    logging.info(f'Processando página {pagina}/{total_paginas}')
    bot.process_page(pagina)
