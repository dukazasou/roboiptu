import pandas as pd
import logging
from selenium import webdriver
import os
import PyPDF2
import time
import wget
from datetime import date


def mover_historico(log):

    # Mover log para histórico

    historico='historico.log'
    novo_historico = 'novo_historico.log'

    with open(historico, 'r') as f1, open(log, 'r') as f2, open(novo_historico, 'w') as fout:

        # Ler linhas do primeiro arquivo
        linhas_historico = f1.readlines()
        # Ler linhas do segundo arquivo
        linhas_log = f2.readlines()

        # Mesclar as linhas
        linhas_mescladas = linhas_historico + linhas_log

        # Ordenar as linhas se necessário
        # linhas_mescladas.sort()

        # Escrever as linhas mescladas no arquivo de saída
        fout.writelines(linhas_mescladas)



    # Substituir o histórico por um novo
    with open(novo_historico, 'r') as arquivo_origem, open(historico, 'w') as arquivo_destino:
        # Ler o conteúdo do arquivo de origem
        conteudo_origem = arquivo_origem.read()

        # Escrever o conteúdo no arquivo de destino
        arquivo_destino.write(conteudo_origem)

    # Apagar o historico novo(temporário)
    os.remove('novo_historico.log')

    # Limpar robo_enfiteutica.log
    with open(log, 'w') as arquivo:
        # Sobrescrever o arquivo vazio
        arquivo.write("")

def extracao():
    while True:
        # controle de linha para fazer a busca na planilha de inscrições
        linha = 0

        # Controlador das vezes que o número de inscrição veio fora do padrão
        cont = 0

        # Número de vezes que precisou reininicar a página web
        web = 0

        # Número de vezes que foi necessário esperar o site retornar
        num = 0

        # Abrir navegador Chrome
        try:
            navegador = webdriver.Chrome()
            logging.info('Abrir navegador')
        except Exception as erro:
            logging.warning(erro)

        # Ler a planilha com as inscrições
        try:
            planilha = pd.read_excel(r"Inscricoes/planilha.xlsx")
            logging.info(f'Ler a planilha')
        except FileNotFoundError as erro:
            logging.critical(erro)
            logging.info('Fim do programa')
            exit()
        except Exception as erro:
            logging.warning(erro)

        # Ir ao site da prefeitura
        try:
            navegador.get('https://daminternet.rio.rj.gov.br/GuiaPagamento/EmitirRegularizacao')
            logging.info('Ir ao site Carioca Digital')
        except Exception as erro:
            logging.critical(erro)

        while True:
            while True:
                # Copiar número de inscrição
                try:
                    inscri = int(planilha['INSCRIÇÃO'][linha])
                    logging.info(f'Inscrição:  {inscri}')

                except Exception as erro:
                    logging.warning(erro)
                    cont = cont + 1
                    if cont == 2:
                        logging.info(f'Fim da planilha')
                        return
                    linha = linha + 1


                # Selecionar inscrição imobiliária
                try:
                    navegador.find_element('xpath',
                                           '//*[@id="divMain"]/div[3]/div/div/div[2]/div[2]/div/div/div/div/div[1]/div/div/div/div[1]/div/select').send_keys(
                        'Inscrição')
                    logging.info('Selecionar inscrição imobiliária')
                except Exception as erro:

                    logging.warning(erro)
                    navegador.get('https://daminternet.rio.rj.gov.br/GuiaPagamento/EmitirRegularizacao')
                    logging.info('Reiniciar o site Carioca Digital')
                    web = web + 1
                    break

                # Inserir número de inscrição
                try:
                    navegador.find_element('xpath', '//*[@id="tbInscricaoImobiliaria"]').send_keys(inscri)
                    logging.info('Inserir número de inscrição')
                except Exception as erro:
                    logging.warning(erro)
                    navegador.get('https://daminternet.rio.rj.gov.br/GuiaPagamento/EmitirRegularizacao')
                    logging.info('Reiniciar o site Carioca Digital')
                    web = web + 1
                    break


                # Seguir para consultar
                try:
                    navegador.find_element('xpath',
                                           '//*[@id="divMain"]/div[3]/div/div/div[2]/div[2]/div/div/div/div/div[1]/div/div/div/div[3]/div/button').click()
                    logging.info('Clicar no botão consultar')
                except Exception as erro:
                    logging.warning(erro)
                    navegador.get('https://daminternet.rio.rj.gov.br/GuiaPagamento/EmitirRegularizacao')
                    logging.info('Reiniciar o site Carioca Digital')
                    web = web + 1
                    break
                time.sleep(5)

                # Descer a página
                try:
                    navegador.execute_script('window.scrollBy(0, 500)')
                    logging.info('Descer a página')
                except Exception as erro:
                    logging.critical(erro)

                # Verificar se há registros para o número de inscrição informado
                try:
                    navegador.find_element('xpath', '//*[@id="fancymodal-1"]/div[2]/div[2]/table/tbody/tr[1]/td[2]')
                    logging.warning(f'Não há registros para o número de inscrição: {inscri}')
                    linha = linha + 1
                    break
                except:
                    logging.info(f'Há registros para o número de inscrição: {inscri}')

                # Selecionar contribuinte
                try:
                    navegador.find_element('xpath',
                                           '//*[@id="divMain"]/div[3]/div/div/div[2]/div[2]/div/div/div/div/div[4]/table/tbody/tr/td[1]/input').click()
                    logging.info('Selecionar contribuinte')
                except Exception as erro:
                    logging.warning(erro)
                    navegador.get('https://daminternet.rio.rj.gov.br/GuiaPagamento/EmitirRegularizacao')
                    logging.info('Reiniciar o site Carioca Digital')
                    web = web + 1
                    break

                # Consultar parcela em atraso
                try:
                    navegador.find_element('xpath',
                                           '//*[@id="divMain"]/div[3]/div/div/div[2]/div[2]/div/div/div/div/table/tbody/tr/td/input').click()
                    logging.info('Clicar botão Consultar')
                except Exception as erro:
                    logging.warning(erro)
                    navegador.get('https://daminternet.rio.rj.gov.br/GuiaPagamento/EmitirRegularizacao')
                    logging.info('Reiniciar o site Carioca Digital')
                    web = web + 1
                    break

                ##################################################################################
                # Copiar a data de vencimento
                try:
                    data = str(planilha['VENCIMENTO'][linha])
                    if len(data) == 3:
                        logging.critical('Data de vencimento vazio')
                        return
                    logging.info(f'Data de vencimento:  {data}')
                except Exception as erro:
                    logging.warning(erro)
                    break
                time.sleep(5)

                # Selecionar a data de vencimento
                try:
                    navegador.find_element('xpath', '//*[@id="dataVencimento"]').send_keys(data)
                    logging.info('Selecionar a data de vencimento')
                except Exception as erro:
                    logging.warning(erro)
                    navegador.get('https://daminternet.rio.rj.gov.br/GuiaPagamento/EmitirRegularizacao')
                    logging.info('Reiniciar o site Carioca Digital')
                    web = web + 1
                    break
                time.sleep(5)

                # Selecionar a primeira parcela
                try:
                    navegador.find_element('xpath',
                                           '//*[@id="divMain"]/div[3]/div/div/div[2]/div[2]/div/table/tbody/tr/td/div/table/tbody/tr[1]/td[1]/label/input').click()
                    logging.info('Escolher por padrão a primeira parcela')
                except Exception as erro:
                    logging.warning(erro)
                    navegador.get('https://daminternet.rio.rj.gov.br/GuiaPagamento/EmitirRegularizacao')
                    logging.info('Reiniciar o site Carioca Digital')
                    web = web + 1
                    break

                # Descer a página
                try:
                    navegador.execute_script('window.scrollBy(0, 500)')
                    logging.info('Scroll na página')
                except Exception as erro:
                    logging.critical(erro)

                # Emitir guia
                try:
                    navegador.find_element('xpath',
                                           '//*[@id="divMain"]/div[3]/div/div/div[2]/div[2]/div/button').click()
                    logging.info('Apertar botão emitir guia')
                except Exception as erro:
                    logging.warning(erro)
                    navegador.get('https://daminternet.rio.rj.gov.br/GuiaPagamento/EmitirRegularizacao')
                    logging.info('Reiniciar o site Carioca Digital')
                    web = web + 1
                    break
                time.sleep(15)

                # Copiar nome da janela Carioca Digital
                try:
                    cariocaDigital = navegador.window_handles[0]
                    logging.info('Copiar nome da janela Carioca Digital')
                except Exception as erro:
                    logging.critical(erro)

                # Copiar nome da janela do pdf
                try:
                    pdf = navegador.window_handles[1]
                    logging.info('Copiar nome da janela do pdf')
                except Exception as erro:
                    logging.critical(erro)

                # Selecionar o pdf para navegar
                try:
                    navegador.switch_to.window(pdf)
                    logging.info('Escolher janela PDF')
                except Exception as erro:
                    logging.critical(erro)

                # Copiar link do pdf
                try:
                    url = navegador.current_url
                    logging.info(f'Link do pdf:  {url}')
                except Exception as erro:
                    logging.critical(erro)

                # Copiar Empreendimento da planilha
                try:
                    empreendimento = str(planilha['EMPREENDIMENTO'][linha])
                    logging.info(f'Nome do empreendimento:  {empreendimento}')
                except Exception as erro:
                    logging.critical(erro)


                # Copiar bloco da planilha
                try:
                    bloco = str(planilha['BLOCO'][linha])
                    logging.info(f'Número do bloco:  {bloco}')
                except Exception as erro:
                    logging.critical(erro)


                # Copiar unidade da planilha
                try:
                    unidade = str(planilha['UNIDADE'][linha])
                    logging.info(f'Número da unidade:  {unidade}')
                except Exception as erro:
                    logging.critical(erro)

                # Criar pasta para salvar
                # Verifica se a pasta já existe antes de criar
                try:
                    if not os.path.exists('Boletos' + '\\' + empreendimento):
                        logging.info('Precisa criar pasta em boletos')
                        try:
                            os.mkdir('Boletos' + '\\' + empreendimento)
                            logging.info("Pasta criada")
                        except Exception as erro:
                            logging.critical(erro)
                    else:
                        logging.info('Não precisa criar pasta, já existe')
                except:
                    logging.critical('Erro ao verificar se pasta já estava criada e criar caso não')


                # Criar pasta da data que fica dentro da pasta do empreendimento
                # Pegar a data de hoje
                data_hoje = date.today()
                data_hoje = data_hoje.strftime("%d-%m-%Y")
                # Verifica se a pasta já existe antes de criar
                try:
                    if not os.path.exists('Boletos' + '\\' + empreendimento + '\\'+ data_hoje):
                        logging.info('Precisa criar pasta em boletos')
                        try:
                            os.mkdir('Boletos' + '\\' + empreendimento+ '\\'+ data_hoje)
                            logging.info("Pasta criada")
                        except Exception as erro:
                            logging.critical(erro)
                    else:
                        logging.info('Não precisa criar pasta, já existe')
                except:
                    logging.critical('Erro ao verificar se pasta já estava criada e criar caso não')

                # Criar nome salvo do arquivo
                try:
                    nome_arquivo = 'Boletos' + '\\' + empreendimento + '\\' + data_hoje + '\\' + 'BL ' + bloco + ' - ' + unidade + '.PDF'
                    logging.info(f'Boleto gerado:  {nome_arquivo}')
                except Exception as erro:
                    logging.critical(erro)


                # Baixar arquivo com nome montado
                try:
                    wget.download(url, nome_arquivo)
                    logging.info('Baixar arquivo')
                except Exception as erro:
                    logging.critical(erro)


                # Abrir documento PDF
                try:
                    arquivo_pdf = open(nome_arquivo, 'rb')
                    logging.info("Abrir guia baixada")
                except Exception as erro:
                    logging.critical(erro)


                # Faz a leitura usando a biblioteca pypdf2
                try:
                    ler_pdf = PyPDF2.PdfReader(arquivo_pdf)
                    logging.info("Ler guia baixada")
                except Exception as erro:
                    logging.critical(erro)


                ##################################################################################################################################################

                # ler todas as páginas
                try:
                    pagina = ler_pdf.pages[-1]
                    logging.info("Ler todas as páginas")
                except Exception as erro:
                    logging.critical(erro)


                # extrai apenas o texto
                try:
                    conteudo_da_pagina = pagina.extract_text()
                    logging.info("Extrair texto do PDF")
                except Exception as erro:
                    logging.critical(erro)


                # faz a junção das linhas
                try:
                    parsed = ''.join(conteudo_da_pagina)
                    logging.info("Concatenar as linhas")
                except Exception as erro:
                    logging.critical(erro)

                    # Encontrar endereço inicial do valor a partir do R$
                    try:
                        endereco_ini_valor = parsed.find('R$')
                        logging.info("Encontrar palavra endereço inicial do valor a partir do R$")
                    except Exception as erro:
                        logging.critical(erro)
                        enviar_email_admim(log)

                    # Encontrar a próxima quebra de linha a partir do R$
                    try:
                        endereco_fim_valor = parsed.find('\n', endereco_ini_valor)
                        logging.info('Encontrar a próxima quebra de linha a partir do R$')
                    except Exception as erro:
                        logging.critical(erro)
                        enviar_email_admim(log)

                    # Pegar o valor
                    valor = parsed[endereco_ini_valor:endereco_fim_valor].replace('R$', "")

                    # Localizar célula respectiva ao valor da linha selecionada e preencher a celula com valor
                    try:
                        planilha.loc[linha, 'VALOR'] = valor
                        logging.info(
                            'Localizar célula respectiva ao valor da linha selecionada e preencher a celula com valor')
                    except Exception as erro:
                        logging.critical(erro)
                        enviar_email_admim(log)


                # Salvar planilha
                try:
                    nome_planilha = r"Inscricoes\planilha_final.xlsx"
                    planilha.to_excel(nome_planilha)
                    logging.info(f'Salvar planilha: {nome_planilha}')
                except Exception as erro:
                    logging.critical(erro)


                # Fechar pdf
                try:
                    navegador.close()
                    logging.info('Fechar aba do navegador pdf')
                except Exception as erro:
                    logging.critical(erro)


                ################################################################################################################

                # Escolher janela carioca digital
                try:
                    navegador.switch_to.window(cariocaDigital)
                    logging.info('Escolher janela carioca digital')
                except Exception as erro:
                    logging.critical(erro)


                # Avançar para próxima linha da planilha
                try:
                    linha = linha + 1
                    logging.info('Avançar para próxima linha da planilha')
                except Exception as erro:
                    logging.critical(erro)


                if web >= 5:
                    logging.info('Número de vezes de atualização web excedida')

                    # Fechar Carioca Digital
                    try:
                        navegador.close()
                        logging.info('Fechar aba do navegador Carioca Digital')
                    except Exception as erro:
                        logging.critical(erro)


                    # Tempo de espera para retorno do site
                    logging.info('Tempo de espera para retorno do site')
                    time.sleep(20)
                    # Reinicar a contagem de retornos do site
                    web = 0
                    # Número de vezes que o site foi colocado em espera
                    num = num + 1

                    # Abrir navegador Chrome
                    try:
                        navegador = webdriver.Chrome()
                        logging.info('Abrir navegador')
                    except Exception as erro:
                        logging.warning(erro)

                    # Ir ao site da prefeitura
                    try:
                        navegador.get('https://daminternet.rio.rj.gov.br/GuiaPagamento/EmitirRegularizacao')
                        logging.info('Ir ao site Carioca Digital')
                        break
                    except Exception as erro:
                        logging.warning(erro)
                if num >= 10:
                    logging.critical('Número máximo de espera do site atingido, FECHAR PROGRAMA')
                    return




#Configuração do LOG

logging.basicConfig(level=logging.INFO, filename="robo.log", format="%(asctime)s - %(levelname)s - %(message)s")

log = r"robo.log"


mover_historico(log)
logging.info('-'*20+'INICIAR PROGRAMA'+'-'*20)
logging.info('Ir para extração')
extracao()

logging.info('Fim do programa')
