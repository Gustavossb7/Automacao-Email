import time
import random
import string
import math
from selenium import webdriver
import pyautogui




senha1= '0codigoNietzsche'
autenticacao1 = 3301

def leiaSenha(msg):
	ok = False
	while True:
		senha = str(input(msg))
		if senha == senha1 :
			ok = True
		else:
			print('\033[0;31mERRO! Digite a senha válida\033[m')
		if ok:
			break
	return senha

def leiaAutenticacao(msg):
	ok = False
	while True:
		autenticacao = int(input(msg))
		if autenticacao == autenticacao1 :
			ok = True
		else:
			print('\033[0;31mERRO! Digite a autenticação válida\033[m')
		if ok:
			break
	return autenticacao

def leiaOpcao(msg):
	ok = False
	while True:
		opcao = int(input(msg))
		if opcao == 1 or opcao == 2 or opcao == 3 or opcao == 4 or opcao == 5 or opcao == 6:
			ok = True
		else:
			print('\033[0;31mERRO! Digite uma opção válida\033[m')
		if ok:
			break
	return opcao

def leiaCriptografia(msg):
	ok = False
	while True:
		criptografar = str(input(msg))
		if criptografar == 'sim' or  criptografar == 'nao' or criptografar == 'não':
			ok = True
		else:
			print('\033[0;31mERRO! Digite uma opção válida\033[m')
		if ok:
			break
	return criptografar

def leiaNcarteira(msg):
	ok = False
	while True:
		ncarteira = int(input(msg))
		if ncarteira == 1 or ncarteira == 2 or ncarteira == 3 or ncarteira == 4 or ncarteira == 5 or ncarteira == 6 :
			ok = True
		else:
			print('\033[0;31mERRO! Digite uma opção válida\033[m')
		if ok:
			break
	return ncarteira

def criptografia(msg):
	tradutor=""
	for letra in msg:
		if letra in "Aa" : tradutor = tradutor + "H"
		elif letra in "Bb" : tradutor = tradutor + "H"
		elif letra in "Cc": tradutor = tradutor + "?"
		elif letra in "Dd": tradutor = tradutor + "/"
		elif letra in "Ee": tradutor = tradutor + "z"
		elif letra in "Ff": tradutor = tradutor + "g"
		elif letra in "Gg": tradutor = tradutor + "A"
		elif letra in "Hh": tradutor = tradutor + "L"
		elif letra in "Ii": tradutor = tradutor + "<H>"
		elif letra in "Jj": tradutor = tradutor + "-"
		elif letra in "Kk": tradutor = tradutor + "´"
		elif letra in "Ll": tradutor = tradutor + "J"
		elif letra in "Mm": tradutor = tradutor + "j"
		elif letra in "Nn": tradutor = tradutor + "wW"
		elif letra in "Oo": tradutor = tradutor + "9!"
		elif letra in "Pp": tradutor = tradutor + "P"
		elif letra in "Qq": tradutor = tradutor + "r"
		elif letra in "Rr": tradutor = tradutor + "r"
		elif letra in "Ss": tradutor = tradutor + "r"
		elif letra in "Tt": tradutor = tradutor + "ah"
		elif letra in "Uu": tradutor = tradutor + "["
		elif letra in "Vv": tradutor = tradutor + "+"
		elif letra in "Ww": tradutor = tradutor + "ll"
		elif letra in "Xx": tradutor = tradutor + "D"
		elif letra in "Yy": tradutor = tradutor + "Y"
		elif letra in "Zz": tradutor = tradutor + "i"
		elif letra in "1":tradutor = tradutor + "gI-"
		elif letra in "2": tradutor = tradutor + "Ç"
		elif letra in "3": tradutor = tradutor + "]"
		elif letra in "4": tradutor = tradutor + "1"
		elif letra in "5": tradutor = tradutor + "_"
		elif letra in "6": tradutor = tradutor + ":)"
		elif letra in "7": tradutor = tradutor + ":("
		elif letra in "8": tradutor = tradutor + ":):"
		elif letra in "9": tradutor = tradutor + "^"
		else: tradutor = tradutor + letra
	return tradutor




print('')
print('Access Matheus')
print('')
senha = leiaSenha('Digite sua senha: ')
print('')
print('Etapa de verificação credencial de segurança concluida com sucesso.')
print('Por favor verifique a credencial de autentificação de triplo fator.')
print('')
autenticacao = leiaAutenticacao('Digite sua autenticação: ')
print('')
print('Etapa de verificação credencial de autenticação concluida com sucesso.')
print('')
print('')
print('')
print('Bem vindo Matheus')
print('')

saldo1 = float(276.51265865185)  # paxg
saldo2 = float(293.79469940991)  # paxg
saldo3 = float(311.67412551236)  # paxg
saldo4 = float(328.35878201318)  # paxg
saldo5 = float(508_228.49662)  # USDT
saldo6 = float(446_625.37828)  # USDT


while True:
	print('')
	print('')
	opcao = leiaOpcao('Digite o número da opção que deseja: ')

	# rede
	if opcao == 1:
			print(' ')
			print('''Adaptador Ethernet Ethernet:

		   Sufixo DNS específico de conexão. . . . . . :
		   Endereço IPv6 . . . . . . . . . . : 2804:7f4:c780:3628:ada2:7b8:6f2:8c25
		   Endereço IPv6 Temporário. . . . . . . . : 2804:7f4:c780:3628:754a:1e0a:1120:728f
		   Endereço IPv6 de link local . . . . . . . . : fe80::ada2:7b8:6f2:8c25%17
		   Endereço IPv4. . . . . . . .  . . . . . . . : 192.168.15.77
		   Máscara de Sub-rede . . . . . . . . . . . . : 255.255.255.0
		   Gateway Padrão. . . . . . . . . . . . . . . : fe80::dac6:78ff:fecf:7320%17
		                                                 192.168.15.1

		Adaptador de Rede sem Fio Wi-Fi:

		   Estado da mídia. . . . . . . . . . . . . .  : mídia desconectada
		   Sufixo DNS específico de conexão. . . . . . :

		Adaptador de Rede sem Fio Conexão Local* 9:

		   Estado da mídia. . . . . . . . . . . . . .  : mídia desconectada
		   Sufixo DNS específico de conexão. . . . . . :

		Adaptador de Rede sem Fio Conexão Local* 10:

		   Estado da mídia. . . . . . . . . . . . . .  : mídia desconectada
		   Sufixo DNS específico de conexão. . . . . . :''')
			print(' ')
			print(' ')
			print(' ')
			print('A rede não possui segurança a nível de criptografia de ponta a ponta, ')
			print(' ')

			criptografar = leiaCriptografia(
				'Deseja criptografar a conexão do programa, assim criando um a canal seguro? ')

			if criptografar == 'sim'.upper().lower():
				print('')
				print(' ')
				print('''JDxK4ZlvptmPYmK3LmYd3QoSFXpHI4FBYOQJ2GBG7LrBz4lOODYq291W7lSnraA-dXOO
					u1JjP8Lsusb5jP6NyYkUFDlCBCs4Rixcx2IKhOB88opiTEBQHyTfNf2XDjDORLlaFjOlLaxr2nxnQ
					juydnq0zRsmF0m83JqQbNkcUmPTwzuVxVQ5MXnHoyjtRlxk0eR3OmKwtmG_Re5q9BYPpZQ7mepTZp
					DJ6fkJMfqYdtS_n_RH43E3Bw_LM3K90qOR''')
				print(' ')
				print(' ')


			if criptografar == 'não'.upper().lower():
				print('')
				print("Não foi possível fazer o login")
			# voltar para o menu

			if criptografar == 'nao'.upper().lower():
				print('')
				print("Não foi possível fazer o login")


	# informações rede
	if opcao == 2:
		print("")
		print("Informações de rede: ")
		print("")
		print("")
		print("")
		print("")
		print("O sistema possui ID de rede classe dhcp e IPv6 permitidas para o adaptador ")
		print("")
		print('''
	Possuimos uma rede padrão que exibe apenas o endereço IP, a máscara de sub-rede e
	o gateway padrão para cada adaptador limitado ao TCP/IP.
	Para Release e Renew, se nenhum nome de adaptador for especificado,
	as concessões de endereços IP para todos os adaptadores limitados
	ao TCP/IP serão liberadas ou renovadas.

	Para Setclassid e Setclassid6, se nenhuma ClassId for especificada, ClassId
	será removida.''')
		time.sleep(1.5)
		print("")
		print("")
		print('''
	Nome do host. . . . . . . . . . . . . . . . : DESKTOP-F2UG63E
	   Sufixo DNS primário . . . . . . . . . . . . :
	   Tipo de nó. . . . . . . . . . . . . . . . . : híbrido
	   Roteamento de IP ativado. . . . . . . . . . : não
	   Proxy WINS ativado. . . . . . . . . . . . . : não
	   Lista de pesquisa de sufixo DNS . . . . . . : virtua.com.br

	Adaptador Ethernet Ethernet:

	   Estado da mídia. . . . . . . . . . . . . .  : mídia desconectada
	   Sufixo DNS específico de conexão. . . . . . :
	   Descrição . . . . . . . . . . . . . . . . . : Realtek PCIe GbE Family Controller
	   Endereço Físico . . . . . . . . . . . . . . : A8-A1-59-79-9F-9F
	   DHCP Habilitado . . . . . . . . . . . . . . : Sim
	   Configuração Automática Habilitada. . . . . : Sim

	Adaptador de Rede sem Fio Conexão Local* 9:

	   Estado da mídia. . . . . . . . . . . . . .  : mídia desconectada
	   Sufixo DNS específico de conexão. . . . . . :
	   Descrição . . . . . . . . . . . . . . . . . : Microsoft Wi-Fi Direct Virtual Adapter
	   Endereço Físico . . . . . . . . . . . . . . : 2A-EE-52-BC-E0-C1
	   DHCP Habilitado . . . . . . . . . . . . . . : Sim
	   Configuração Automática Habilitada. . . . . : Sim''')
		time.sleep(2)
		print('''
	Adaptador de Rede sem Fio Conexão Local* 10:

	   Estado da mídia. . . . . . . . . . . . . .  : mídia desconectada
	   Sufixo DNS específico de conexão. . . . . . :
	   Descrição . . . . . . . . . . . . . . . . . : Microsoft Wi-Fi Direct Virtual Adapter #2
	   Endereço Físico . . . . . . . . . . . . . . : 28-EE-52-BC-E0-C1
	   DHCP Habilitado . . . . . . . . . . . . . . : Sim
	   Configuração Automática Habilitada. . . . . : Sim

	Adaptador de Rede sem Fio Wi-Fi:

	   Sufixo DNS específico de conexão. . . . . . : virtua.com.br
	   Descrição . . . . . . . . . . . . . . . . . : TP-Link Wireless USB Adapter
	   Endereço Físico . . . . . . . . . . . . . . : 28-EE-52-BC-E0-C1
	   DHCP Habilitado . . . . . . . . . . . . . . : Sim
	   Configuração Automática Habilitada. . . . . : Sim
	   Endereço IPv6 . . . . . . . . . . : 2804:14d:4cd1:84a0::504a(Preferencial)
	   Concessão Obtida. . . . . . . . . . . . . . : quinta-feira, 11 de agosto de 2022 15:55:36
	   Concessão Expira. . . . . . . . . . . . . . : terça-feira, 16 de agosto de 2022 16:10:20
	   Endereço IPv6 de link local . . . . . . . . : fe80::858:6458:4e62:469%19(Preferencial)
	   Endereço IPv4. . . . . . . .  . . . . . . . : 192.168.0.219(Preferencial)
	   Máscara de Sub-rede . . . . . . . . . . . . : 255.255.255.0
	   Concessão Obtida. . . . . . . . . . . . . . : quinta-feira, 11 de agosto de 2022 15:55:41
	   Concessão Expira. . . . . . . . . . . . . . : sábado, 13 de agosto de 2022 04:10:31
	   Gateway Padrão. . . . . . . . . . . . . . . : fe80::ce75:e2ff:fed8:4fca%19
	                                                 192.168.0.1
	   Servidor DHCP . . . . . . . . . . . . . . . : 192.168.0.1
	   IAID de DHCPv6. . . . . . . . . . . . . . . : 405335634
	   DUID de Cliente DHCPv6. . . . . . . . . . . : 00-01-00-01-2A-76-4F-9E-A8-A1-59-79-9F-9F
	   Servidores DNS. . . . . . . . . . . . . . . : 2804:14d:1:0:181:213:132:4
	                                                 2804:14d:1:0:181:213:132:5
	                                                 181.213.132.4
	                                                 181.213.132.5
	   NetBIOS em Tcpip. . . . . . . . . . . . . . : Habilitado''')
		print("")
		print("")


	# cotação
	if opcao == 3:
			print("")
			print("")
			print("")
			print("")
			print("A aplicação está consultando as cotações atuais...")
			print("")

			navegador = webdriver.Chrome()
			print("")
			navegador.get("https://coinmarketcap.com/pt-br/currencies/tether/")

			print("")
			print("")
			cotacaoUSDT = navegador.find_element("xpath",
												 '//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[2]/div/div[2]/div[1]/div/span')
			precoUSDT = navegador.find_element("xpath",
											   '// *[ @ id = "__next"] / div / div[1] / div[2] / div / div[1] / div[2] / div / div[2] / h1')
			marketcapUSDT = navegador.find_element("xpath",
												   '//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[2]/div/div[3]/div[1]/div[1]/div/div[2]/div')
			fullydilutedmarketcapUSDT = navegador.find_element("xpath",
															   '//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[2]/div/div[3]/div[1]/div[2]/div/div[2]/div')
			volume24hUSDT = navegador.find_element("xpath",
												   '//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[2]/div/div[3]/div[1]/div[3]/div[1]/div[2]/div')
			print(precoUSDT.text)
			print(cotacaoUSDT.text)
			print("")
			print("Market Cap", '         ', "Fully Diluted Market Cap", '     ', "Volume 24h")
			print(marketcapUSDT.text, '     ', fullydilutedmarketcapUSDT.text, '     ', volume24hUSDT.text)
			print("")
			print("")
			print("")

			print('------------------------------')
			print("")
			print("")

			navegador.get("https://coinmarketcap.com/pt-br/currencies/pax-gold/")
			cotacaoPAXG = navegador.find_element("xpath",
												 '//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[2]/div/div[2]/div[1]/div/span')
			precoPAXG = navegador.find_element("xpath",
											   '// *[ @ id = "__next"] / div / div[1] / div[2] / div / div[1] / div[2] / div / div[2] / h1')
			marketcapPAXG = navegador.find_element("xpath",
												   '//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[2]/div/div[3]/div[1]/div[1]/div/div[2]/div')
			fullydilutedmarketcapPAXG = navegador.find_element("xpath",
															   '//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[2]/div/div[3]/div[1]/div[2]/div/div[2]/div')
			volume24hPAXG = navegador.find_element("xpath",
												   '//*[@id="__next"]/div/div[1]/div[2]/div/div[1]/div[2]/div/div[3]/div[1]/div[3]/div[1]/div[2]/div')

			print('Preço de PAX Gold (PAXG)')
			print(cotacaoPAXG.text)
			print("")
			print("Market Cap", '         ', "Fully Diluted Market Cap", '     ', "Volume 24h")
			print(marketcapPAXG.text, '     ', fullydilutedmarketcapPAXG.text, '            ', volume24hPAXG.text)
			print("")
			print("")
			print("")


	# carteiras
	if opcao == 4:
			print('''	Carteiras dísponíveis:
			Carteira R1
			Carteira R2
			Carteira R3
			Carteira R4
			Carteira I1
			Carteira FTX''')
			print(' ')
			opcao1 = leiaCriptografia('Deseja inspecionar o circuito interno? ')
			if opcao1 == 'não'.upper().lower() or opcao1 == 'nao'.upper().lower():
				print('')
			if opcao1 == 'sim'.upper().lower():

				print('')
				ncarteira = leiaNcarteira('Digite o número da carteira: ')
				print('')

				total = float(15_896_777.39)
				# real1 = float(2_543_484.38)
				# real2 = float(2_702_452.1563)
				# real3 = float(2_861_419.9302)
				# real4 = float(3_020_387.7041)
				# real5 = float(2_622_968.26)
				# real6 = float(2_305_032.72)



				real1 = float(saldo1 * 9097.38)
				real2 = float(saldo2 * 9097.38)
				real3 = float(saldo3 * 9097.38)
				real4 = float(saldo4 * 9097.38)
				real5 = float(saldo5 * 5.07)
				real6 = float(saldo6 * 5.07)

				la1 = ''.join(random.SystemRandom().choices(string.ascii_letters, k=1))
				la2 = ''.join(random.SystemRandom().choices(string.ascii_letters, k=1))
				la3 = ''.join(random.SystemRandom().choices(string.ascii_letters, k=1))

				na1 = random.choice([1, 2, 3, 4, 5, 6, 7, 8, 9])
				na2 = random.choice([1, 2, 3, 4, 5, 6, 7, 8, 9])
				na3 = random.choice([1, 2, 3, 4, 5, 6, 7, 8, 9])
				na4 = random.choice([1, 2, 3, 4, 5, 6, 7, 8, 9])
				na5 = random.choice([1, 2, 3, 4, 5, 6, 7, 8, 9])
				na6 = random.choice([1, 2, 3, 4, 5, 6, 7, 8, 9])
				na7 = random.choice([1, 2, 3, 4, 5, 6, 7, 8, 9])
				na8 = random.choice([1, 2, 3, 4, 5, 6, 7, 8, 9])

				ID1 = (na1,na2,na3,la1,na4,na5,na6,na7,na8,"-",la2,la3)
				ID2 = (na1, na2, na3, la1, na4, na5, na6, na7, na8, "-", la2, la3)
				ID3 = (na1, na2, na3, la1, na4, na5, na6, na7, na8, "-", la2, la3)
				ID4 = (na1, na2, na3, la1, na4, na5, na6, na7, na8, "-", la2, la3)
				ID5 = (na1, na2, na3, la1, na4, na5, na6, na7, na8, "-", la2, la3)
				ID6 = (na1, na2, na3, la1, na4, na5, na6, na7, na8, "-", la2, la3)

				ta1 = 100 * random.random()
				ta1 = math.ceil(ta1)
				while ta1 < 7 and ta1 > 47:
					ta1 = 100 * random.random()
					ta1 = math.ceil(ta1)

				ta2 = 100 * random.random()
				ta2 = math.ceil(ta2)
				while ta2 < 7 and ta2 > 53:
					ta2 = 100 * random.random()
					ta2 = math.ceil(ta2)

				ta3 = 160-(ta1+ta2)

				al1 = round((100 * real1) / total,3)
				al2 = round((100 * real2) / total,3)
				al3 = round((100 * real3) / total,3)
				al4 = round((100 * real4) / total,3)
				al5 = round((100 * real5) / total,3)
				al6 = round((100 * real6) / total,3)


				if ncarteira == 1:
					print(
						'Criptomoeda: Paxos Gold (PAXG), Quantidade: {}, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%'.format(
							saldo1, al1))
					print('')
					transfer = str(input('Deseja movimentar o circuito interno? '))
					print('')
					if transfer == 'não'.upper().lower() or transfer =='nao'.upper().lower():
						print('')
					if transfer == 'sim'.upper().lower():
						print('')
						valor = float(input('Qualifique a tranferencia: '))
						print('')
						destino = leiaNcarteira('Qual a carteira de destino? (digite o número da carteira): ')
						# taxa = random.choice(
						# 	[0.001562, 0.002604, 0.002322, 0.002298, 0.002512, 0.002233, 0.001666, 0.001732, 0.001878,
						# 	 0.001904,
						# 	 0.002043, 0.002179, 0.002417, 0.002471, 0.002438, 0.001763, 0.001589, 0.001622, 0.001927,
						# 	 0.001856])
						# taxa2=taxa*10
						taxaPaxg = random.choice(
							[0.00098869808, 0.000860438248, 0.000332286556, 0.000229802845, 0.000451243811,
							 0.000223358254, 0.001666435356, 0.00192716402, 0.005312830272, 0.007470163288])

						taxaPaxg2 = random.choice(
							[0.014144054453, 0.017594981639, 0.019175284100, 0.025124769641, 0.026274362743,
							 0.021119023841, 0.022372391771, 0.039284750228, 0.030672705062, 0.037388547043])

						taxaUsdt = random.choice(
							[1.77289, 3.86556, 4.29802, 11.11901, 7.01632, 9.144054, 6.00988, 1.39284, 8.85470,
							 13.39448])

						taxaUsdt2 = random.choice(
							[25.41013, 36.41005, 63.033228, 32.902384, 40.40544, 55.07471, 61.91639, 27.45124, 49.73885,
							 67.16942])
						print('')
						if destino == 5 or destino == 6:
							print(
								'''	
                            ERC-20 Transaction: Tax pending.
                            Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                            Conversão será feita automaticamente via protocolo crosschain (1 INCH NETWORK).
                            ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format( na1, na2, na3, la1, na4, na5, na6, na7, na8,
																  la2, la3))
							time.sleep(3)
							confirmar = leiaCriptografia('''
							O valor de taxa de é {} PAXG.
							Digite sim, para confirmar:'''.format(taxaPaxg2))

							if confirmar == 'sim'.upper().lower().strip():
								saldo1 = saldo1 - valor
								if destino == 1:
									saldo1 = saldo1 + valor - taxaPaxg
									real1 = (saldo1 * 9097.38)
									al1 = round((100 * real1) / total, 3)
								if destino == 2:
									saldo2 = saldo2 + valor - taxaPaxg
									real2 = (saldo2 * 9097.38)
									al2 = round((100 * real2) / total, 3)
								if destino == 3:
									saldo3 = saldo3 + valor - taxaPaxg
									real3 = (saldo3 * 9097.38)
									al3 = round((100 * real3) / total, 3)
								if destino == 4:
									saldo4 = saldo4 + valor - taxaPaxg
									real4 = (saldo4 * 9097.38)
									al4 = round((100 * real4) / total, 3)
								if destino == 5:
									saldo5 = saldo5 + (valor * 1794) - taxaPaxg2
									real5 = (saldo5 * 5.07)
									saldo5 = round(saldo5, 5)
									al5 = round((100 * real5) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 6:
									saldo6 = saldo6 + (valor * 1794) - taxaPaxg2
									real6 = (saldo6 * 5.07)
									saldo6 = round(saldo6, 5)
									al6 = round((100 * real6) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)

								for c in range(1, ta1):
									print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta2):
									print('Confirmação de rede pendende (1/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta3):
									print('Confirmação de rede pendende (2/3): PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								print('Confirmação de rede pendende (3/3): PROTOCOLO -',
									  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
								if destino == 1:
									print(
										'O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo1, al1, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 2:
									print(
										'O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo2, al2, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 3:
									print(
										'O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo3, al3, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 4:
									print(
										'O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo4, al4, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 5:
									print(
										'O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo5, al5, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 6:
										print(
											'O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo6, al6, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))

						else:
							print(
								'''	
                            ERC-20 Transaction: Tax pending.
                            Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                            ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format(na1, na2, na3, la1, na4, na5, na6, na7, na8, la2,la3))
							time.sleep(3)
							confirmar = leiaCriptografia('''
							O valor de taxa de é {} PAXG.
							Digite sim, para confirmar:'''.format(taxaPaxg))
							if confirmar == 'sim'.upper().lower().strip():
								saldo1 = saldo1 -valor
								if destino == 1:
									saldo1 = saldo1 + valor - taxaPaxg
									real1 = (saldo1 * 9097.38)
									al1 = round((100 * real1) / total, 3)
								if destino == 2:
									saldo2 = saldo2 + valor - taxaPaxg
									real2 = (saldo2 * 9097.38)
									al2 = round((100 * real2) / total,3)
								if destino == 3:
									saldo3 = saldo3 + valor - taxaPaxg
									real3 = (saldo3 * 9097.38)
									al3 = round((100 * real3) / total,3)
								if destino == 4:
									saldo4 = saldo4 + valor - taxaPaxg
									real4 = (saldo4 * 9097.38)
									al4 = round((100 * real4) / total,3)
								if destino == 5:
									saldo5 = saldo5 + (valor*1794) - taxaPaxg2
									real5 = (saldo5 * 5.07)
									saldo5 = round(saldo5, 5)
									al5 = round((100 * real5) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 6:
									saldo6 = saldo6 + (valor*1794) - taxaPaxg2
									real6 = (saldo6 * 5.07)
									saldo6 = round(saldo6, 5)
									al6 = round((100 * real6) / total,3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)

								for c in range(1,ta1):
									print('Confirmação de rede pendende (0/3):  PROTOCOLO -',''.join(random.SystemRandom().choices(string.ascii_letters, k = 7)))
									time.sleep(1.5)
								for c in range(1,ta2):
									print('Confirmação de rede pendende (1/3):  PROTOCOLO -',''.join(random.SystemRandom().choices(string.ascii_letters, k = 7)))
									time.sleep(1.5)
								for c in range(1,ta3):
									print('Confirmação de rede pendende (2/3): PROTOCOLO -',''.join(random.SystemRandom().choices(string.ascii_letters, k = 7)))
									time.sleep(1.5)
								print('Confirmação de rede pendende (3/3): PROTOCOLO -',''.join(random.SystemRandom().choices(string.ascii_letters, k = 7)))

								if destino == 1:
									print('O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(saldo1, al1,na1,na2,na3,la1,na4,na5,na6,na7,na8,la2,la3))
								if destino == 2:
									print('O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(saldo2, al2,na1,na2,na3,la1,na4,na5,na6,na7,na8,la2,la3))
								if destino == 3:
									print('O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(saldo3, al3,na1,na2,na3,la1,na4,na5,na6,na7,na8,la2,la3))
								if destino == 4:
									print('O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(saldo4, al4,na1,na2,na3,la1,na4,na5,na6,na7,na8,la2,la3))
								if destino == 5:
									print('O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(saldo5, al5,na1,na2,na3,la1,na4,na5,na6,na7,na8,la2,la3))
								if destino == 6:
									print('O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(saldo6, al6,na1,na2,na3,la1,na4,na5,na6,na7,na8,la2,la3))


					else:
						print('Processo Finalizado')

				if ncarteira == 2:
					print(
						'Criptomoeda: Paxos Gold (PAXG), Quantidade: {}, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%'.format(
							saldo2, al2))
					print('')
					transfer = str(input('Deseja movimentar o circuito interno? '))
					print('')
					if transfer == 'não'.upper().lower() or transfer == 'nao'.upper().lower():
						print('')
					if transfer == 'sim'.upper().lower():
						print('')
						valor = float(input('Qualifique a tranferencia: '))
						print('')
						destino = leiaNcarteira('Qual a carteira de destino? (digite o número da carteira): ')
						# taxa = random.choice(
						# 	[0.001562, 0.002604, 0.002322, 0.002298, 0.002512, 0.002233, 0.001666, 0.001732, 0.001878,
						# 	 0.001904,
						# 	 0.002043, 0.002179, 0.002417, 0.002471, 0.002438, 0.001763, 0.001589, 0.001622, 0.001927,
						# 	 0.001856])
						# taxa2 = taxa * 10
						taxaPaxg = random.choice(
							[0.00098869808, 0.000860438248, 0.000332286556, 0.000229802845, 0.000451243811,
							 0.000223358254, 0.001666435356, 0.00192716402, 0.005312830272, 0.007470163288])

						taxaPaxg2 = random.choice(
							[0.014144054453, 0.017594981639, 0.019175284100, 0.025124769641, 0.026274362743,
							 0.021119023841, 0.022372391771, 0.039284750228, 0.030672705062, 0.037388547043])

						taxaUsdt = random.choice(
							[1.77289, 3.86556, 4.29802, 11.11901, 7.01632, 9.144054, 6.00988, 1.39284, 8.85470,
							 13.39448])

						taxaUsdt2 = random.choice(
							[25.41013, 36.41005, 63.033228, 32.902384, 40.40544, 55.07471, 61.91639, 27.45124, 49.73885,
							 67.16942])
						print('')
						if destino == 5 or destino == 6:
							print(
								'''	
                            ERC-20 Transaction: Tax pending.
                            Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                            Conversão será feita automaticamente via protocolo crosschain (1 INCH NETWORK).
                            ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format(na1, na2, na3, la1, na4, na5, na6, na7, na8,
																   la2, la3))
							time.sleep(3)
							confirmar = leiaCriptografia('''
							O valor de taxa de é {} PAXG.
							Digite sim, para confirmar:'''.format(taxaPaxg2))
							if confirmar == 'sim'.upper().lower().strip():
								saldo2 = saldo2 - valor
								if destino == 1:
									saldo1 = saldo1 + valor - taxaPaxg
									real1 = float(saldo1 * 9097.38)
									al1 = round((100 * real1) / total, 3)
								if destino == 2:
									saldo2 = saldo2 + valor - taxaPaxg
									real2 = float(saldo2 * 9097.38)
									al2 = round((100 * real2) / total, 3)
								if destino == 3:
									saldo3 = saldo3 + valor - taxaPaxg
									real3 = float(saldo3 * 9097.38)
									al3 = round((100 * real3) / total, 3)
								if destino == 4:
									saldo4 = saldo4 + valor - taxaPaxg
									real4 = float(saldo4 * 9097.38)
									al4 = round((100 * real4) / total, 3)
								if destino == 5:
									saldo5 = saldo5 + (valor * 1794) - taxaPaxg2
									real5 = (saldo5 * 5.07)
									saldo5 = round(saldo5, 5)
									al5 = round((100 * real5) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 6:
									saldo6 = saldo6 + (valor * 1794) - taxaPaxg2
									real6 = (saldo6 * 5.07)
									saldo6 = round(saldo6, 5)
									al6 = round((100 * real6) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)

								for c in range(1, ta1):
									print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta2):
									print('Confirmação de rede pendende (1/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta3):
									print('Confirmação de rede pendende (2/3): PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								print('Confirmação de rede pendende (3/3): PROTOCOLO -',
									  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))

								if destino == 1:
									print(
										'O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo1, al1, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 2:
									print(
										'O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo2, al2, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 3:
									print(
										'O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo3, al3, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 4:
									print(
										'O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo4, al4, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 5:
									print(
										'O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo5, al5, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 6:
									print(
										'O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo6, al6, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))

						else:
							print(
								'''	
                            ERC-20 Transaction: Tax pending.
                            Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                            ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format(na1, na2, na3, la1, na4, na5, na6, na7, na8, la2,
																   la3))
							time.sleep(3)
							confirmar = leiaCriptografia('''
							O valor de taxa de é {} PAXG.
							Digite sim, para confirmar:'''.format(taxaPaxg))
							if confirmar == 'sim'.upper().lower().strip():
								saldo2 = saldo2 - valor
								if destino == 1:
									saldo1 = saldo1 + valor - taxaPaxg
									real1 = float(saldo1 * 9097.38)
									al1 = round((100 * real1) / total, 3)
								if destino == 2:
									saldo2 = saldo2 + valor - taxaPaxg
									real2 = float(saldo2 * 9097.38)
									al2 = round((100 * real2) / total, 3)
								if destino == 3:
									saldo3 = saldo3 + valor - taxaPaxg
									real3 = float(saldo3 * 9097.38)
									al3 = round((100 * real3) / total, 3)
								if destino == 4:
									saldo4 = saldo4 + valor - taxaPaxg
									real4 = float(saldo4 * 9097.38)
									al4 = round((100 * real4) / total, 3)
								if destino == 5:
									saldo5 = saldo5 + (valor * 1794) - taxaPaxg2
									real5 = (saldo5 * 5.07)
									saldo5 = round(saldo5, 5)
									al5 = round((100 * real5) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 6:
									saldo6 = saldo6 + (valor * 1794) - taxaPaxg2
									real6 = (saldo6 * 5.07)
									saldo6 = round(saldo6, 5)
									al6 = round((100 * real6) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)

								for c in range(1, ta1):
									print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta2):
									print('Confirmação de rede pendende (1/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta3):
									print('Confirmação de rede pendende (2/3): PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								print('Confirmação de rede pendende (3/3): PROTOCOLO -',
									  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))

								if destino == 1:
									print(
										'O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo1, al1, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 2:
									print(
										'O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo2, al2, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 3:
									print(
										'O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo3, al3, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 4:
									print(
										'O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo4, al4, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 5:
									print(
										'O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo5, al5, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 6:
									print(
										'O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo6, al6, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))

					else:
						print('Processo Finalizado')

				if ncarteira == 3:
					print(
						'Criptomoeda: Paxos Gold (PAXG), Quantidade: {}, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%'.format(
							saldo3, al3))
					print('')
					transfer = str(input('Deseja movimentar o circuito interno? '))
					print('')
					if transfer == 'não'.upper().lower() or transfer == 'nao'.upper().lower():
						print('')
					if transfer == 'sim'.upper().lower():
						print('')
						valor = float(input('Qualifique a tranferencia: '))
						print('')
						destino = leiaNcarteira('Qual a carteira de destino? (digite o número da carteira): ')
						# taxa = random.choice(
						# 	[0.001562, 0.002604, 0.002322, 0.002298, 0.002512, 0.002233, 0.001666, 0.001732, 0.001878,
						# 	 0.001904,
						# 	 0.002043, 0.002179, 0.002417, 0.002471, 0.002438, 0.001763, 0.001589, 0.001622, 0.001927,
						# 	 0.001856])
						# taxa2 = taxa * 10
						taxaPaxg = random.choice(
							[0.00098869808, 0.000860438248, 0.000332286556, 0.000229802845, 0.000451243811,
							 0.000223358254, 0.001666435356, 0.00192716402, 0.005312830272, 0.007470163288])

						taxaPaxg2 = random.choice(
							[0.014144054453, 0.017594981639, 0.019175284100, 0.025124769641, 0.026274362743,
							 0.021119023841, 0.022372391771, 0.039284750228, 0.030672705062, 0.037388547043])

						taxaUsdt = random.choice(
							[1.77289, 3.86556, 4.29802, 11.11901, 7.01632, 9.144054, 6.00988, 1.39284, 8.85470,
							 13.39448])

						taxaUsdt2 = random.choice(
							[25.41013, 36.41005, 63.033228, 32.902384, 40.40544, 55.07471, 61.91639, 27.45124, 49.73885,
							 67.16942])
						print('')
						if destino == 5 or destino == 6:
							print(
								'''	
                            ERC-20 Transaction: Tax pending.
                            Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                            Conversão será feita automaticamente via protocolo crosschain (1 INCH NETWORK).
                            ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format(na1, na2, na3, la1, na4, na5, na6, na7, na8,
																   la2, la3))
							time.sleep(3)
							confirmar = leiaCriptografia('''
							O valor de taxa de é {} PAXG.
							Digite sim, para confirmar:'''.format(taxaPaxg2))
							if confirmar == 'sim'.upper().lower().strip():
								saldo3 = saldo3 - valor
								if destino == 1:
									saldo1 = saldo1 + valor - taxaPaxg
									real1 = float(saldo1 * 9097.38)
									al1 = round((100 * real1) / total, 3)
								if destino == 2:
									saldo2 = saldo2 + valor - taxaPaxg
									real2 = float(saldo2 * 9097.38)
									al2 = round((100 * real2) / total, 3)
								if destino == 3:
									saldo3 = saldo3 + valor - taxaPaxg
									real3 = float(saldo3 * 9097.38)
									al3 = round((100 * real3) / total, 3)
								if destino == 4:
									saldo4 = saldo4 + valor - taxaPaxg
									real4 = float(saldo4 * 9097.38)
									al4 = round((100 * real4) / total, 3)
								if destino == 5:
									saldo5 = saldo5 + (valor * 1794) - taxaPaxg2
									real5 = (saldo5 * 5.07)
									saldo5 = round(saldo5, 5)
									al5 = round((100 * real5) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 6:
									saldo6 = saldo6 + (valor * 1794) - taxaPaxg2
									real6 = (saldo6 * 5.07)
									saldo6 = round(saldo6, 5)
									al6 = round((100 * real6) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)

								for c in range(1, ta1):
									print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta2):
									print('Confirmação de rede pendende (1/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta3):
									print('Confirmação de rede pendende (2/3): PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								print('Confirmação de rede pendende (3/3): PROTOCOLO -',
									  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))

								if destino == 1:
									print(
										'O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo1, al1, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 2:
									print(
										'O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo2, al2, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 3:
									print(
										'O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo3, al3, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 4:
									print(
										'O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo4, al4, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 5:
									print(
										'O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo5, al5, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 6:
									print(
										'O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo6, al6, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
						else:
									print(
										'''	
                                    ERC-20 Transaction: Tax pending.
                                    Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                                    ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format(na1, na2, na3, la1, na4, na5, na6, na7, na8,la2,la3))
									time.sleep(3)
									confirmar = leiaCriptografia('''
									O valor de taxa de é {} PAXG.
									Digite sim, para confirmar:'''.format(taxaPaxg))
									if confirmar == 'sim'.upper().lower().strip():
										saldo3 = saldo3 - valor
										if destino == 1:
											saldo1 = saldo1 + valor - taxaPaxg
											real1 = float(saldo1 * 9097.38)
											al1 = round((100 * real1) / total, 3)
										if destino == 2:
											saldo2 = saldo2 + valor - taxaPaxg
											real2 = float(saldo2 * 9097.38)
											al2 = round((100 * real2) / total, 3)
										if destino == 3:
											saldo3 = saldo3 + valor - taxaPaxg
											real3 = float(saldo3 * 9097.38)
											al3 = round((100 * real3) / total, 3)
										if destino == 4:
											saldo4 = saldo4 + valor - taxaPaxg
											real4 = float(saldo4 * 9097.38)
											al4 = round((100 * real4) / total, 3)
										if destino == 5:
											saldo5 = saldo5 + (valor * 1794) - taxaPaxg2
											real5 = (saldo5 * 5.07)
											saldo5 = round(saldo5, 5)
											al5 = round((100 * real5) / total, 3)
											for c in range(1, 60):
												print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
													  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
												time.sleep(1.5)
										if destino == 6:
											saldo6 = saldo6 + (valor * 1794) - taxaPaxg2
											real6 = (saldo6 * 5.07)
											saldo6 = round(saldo6, 5)
											al6 = round((100 * real6) / total, 3)
											for c in range(1, 60):
												print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
													  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
												time.sleep(1.5)

									for c in range(1, ta1):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
									for c in range(1, ta2):
										print('Confirmação de rede pendende (1/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
									for c in range(1, ta3):
										print('Confirmação de rede pendende (2/3): PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
									print('Confirmação de rede pendende (3/3): PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))

									if destino == 1:
										print(
											'O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo1, al1, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 2:
										print(
											'O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo2, al2, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 3:
										print(
											'O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo3, al3, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 4:
										print(
											'O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo4, al4, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 5:
										print(
											'O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo5, al5, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 6:
										print(
											'O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo6, al6, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
					else:
						print('Processo Finalizado')

				if ncarteira == 4:
					print(
						'Criptomoeda: Paxos Gold (PAXG), Quantidade: {}, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%'.format(
							saldo4, al4))
					print('')
					transfer = str(input('Deseja movimentar o circuito interno? '))
					print('')
					if transfer == 'não'.upper().lower() or transfer == 'nao'.upper().lower():
						print('')
					if transfer == 'sim'.upper().lower():
						print('')
						valor = float(input('Qualifique a tranferencia: '))
						print('')
						destino = leiaNcarteira('Qual a carteira de destino? (digite o número da carteira): ')
						# taxa = random.choice(
						# 	[0.001562, 0.002604, 0.002322, 0.002298, 0.002512, 0.002233, 0.001666, 0.001732, 0.001878,
						# 	 0.001904,
						# 	 0.002043, 0.002179, 0.002417, 0.002471, 0.002438, 0.001763, 0.001589, 0.001622, 0.001927,
						# 	 0.001856])
						# taxa2 = taxa * 10
						taxaPaxg = random.choice(
							[0.00098869808, 0.000860438248, 0.000332286556, 0.000229802845, 0.000451243811,
							 0.000223358254, 0.001666435356, 0.00192716402, 0.005312830272, 0.007470163288])

						taxaPaxg2 = random.choice(
							[0.014144054453, 0.017594981639, 0.019175284100, 0.025124769641, 0.026274362743,
							 0.021119023841, 0.022372391771, 0.039284750228, 0.030672705062, 0.037388547043])
						print('')
						if destino == 5 or destino == 6:
							print(
								'''	
                            ERC-20 Transaction: Tax pending.
                            Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                            Conversão será feita automaticamente via protocolo crosschain (1 INCH NETWORK).
                            ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format(na1, na2, na3, la1, na4, na5, na6, na7, na8,
																   la2, la3))
							time.sleep(3)
							confirmar = leiaCriptografia('''
							O valor de taxa de é {} PAXG.
							Digite sim, para confirmar:'''.format(taxaPaxg2))
							if confirmar == 'sim'.upper().lower().strip():
								saldo4 = saldo4 - valor
								if destino == 1:
									saldo1 = saldo1 + valor - taxaPaxg
									real1 = float(saldo1 * 9097.38)
									al1 = round((100 * real1) / total, 3)
								if destino == 2:
									saldo2 = saldo2 + valor - taxaPaxg
									real2 = float(saldo2 * 9097.38)
									al2 = round((100 * real2) / total, 3)
								if destino == 3:
									saldo3 = saldo3 + valor - taxaPaxg
									real3 = float(saldo3 * 9097.38)
									al3 = round((100 * real3) / total, 3)
								if destino == 4:
									saldo4 = saldo4 + valor - taxaPaxg
									real4 = float(saldo4 * 9097.38)
									al4 = round((100 * real4) / total, 3)
								if destino == 5:
									saldo5 = saldo5 + (valor * 1794) - taxaPaxg2
									real5 = (saldo5 * 5.07)
									saldo5 = round(saldo5, 5)
									al5 = round((100 * real5) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 6:
									saldo6 = saldo6 + (valor * 1794) - taxaPaxg2
									real6 = (saldo6 * 5.07)
									saldo6 = round(saldo6, 5)
									al6 = round((100 * real6) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)


								for c in range(1, ta1):
									print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta2):
									print('Confirmação de rede pendende (1/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta3):
									print('Confirmação de rede pendende (2/3): PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								print('Confirmação de rede pendende (3/3): PROTOCOLO -',
									  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))

								if destino == 1:
									print(
										'O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo1, al1, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 2:
									print(
										'O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo2, al2, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 3:
									print(
										'O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo3, al3, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 4:
									print(
										'O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo4, al4, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 5:
									print(
										'O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo5, al5, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 6:
									print(
										'O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo6, al6, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
						else:
									print(
										'''	
                                    ERC-20 Transaction: Tax pending.
                                    Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                                    ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format(na1, na2, na3, la1, na4, na5, na6, na7, na8,
																		   la2, la3))
									time.sleep(3)
									confirmar = leiaCriptografia('''
									O valor de taxa de é {} PAXG.
									Digite sim, para confirmar:'''.format(taxaPaxg))
									if confirmar == 'sim'.upper().lower().strip():
										saldo4 = saldo4 - valor
										if destino == 1:
											saldo1 = saldo1 + valor - taxaPaxg
											real1 = float(saldo1 * 9097.38)
											al1 = round((100 * real1) / total, 3)
										if destino == 2:
											saldo2 = saldo2 + valor - taxaPaxg
											real2 = float(saldo2 * 9097.38)
											al2 = round((100 * real2) / total, 3)
										if destino == 3:
											saldo3 = saldo3 + valor - taxaPaxg
											real3 = float(saldo3 * 9097.38)
											al3 = round((100 * real3) / total, 3)
										if destino == 4:
											saldo4 = saldo4 + valor - taxaPaxg
											real4 = float(saldo4 * 9097.38)
											al4 = round((100 * real4) / total, 3)
										if destino == 5:
											saldo5 = saldo5 + (valor * 1794) - taxaPaxg2
											real5 = (saldo5 * 5.07)
											saldo5 = round(saldo5, 5)
											al5 = round((100 * real5) / total, 3)
											for c in range(1, 60):
												print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
													  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
												time.sleep(1.5)
										if destino == 6:
											saldo6 = saldo6 + (valor * 1794) - taxaPaxg2
											real6 = (saldo6 * 5.07)
											saldo6 = round(saldo6, 5)
											al6 = round((100 * real6) / total, 3)
											for c in range(1, 60):
												print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
													  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
												time.sleep(1.5)


											for c in range(1, ta1):
												print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
													  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
												time.sleep(1.5)
											for c in range(1, ta2):
												print('Confirmação de rede pendende (1/3):  PROTOCOLO -',
													  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
												time.sleep(1.5)
											for c in range(1, ta3):
												print('Confirmação de rede pendende (2/3): PROTOCOLO -',
													  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
												time.sleep(1.5)
											print('Confirmação de rede pendende (3/3): PROTOCOLO -',
												  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))

											if destino == 1:
												print(
													'O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
														saldo1, al1, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
											if destino == 2:
												print(
													'O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
														saldo2, al2, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
											if destino == 3:
												print(
													'O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
														saldo3, al3, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
											if destino == 4:
												print(
													'O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
														saldo4, al4, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
											if destino == 5:
												print(
													'O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
														saldo5, al5, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
											if destino == 6:
												print(
													'O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
														saldo6, al6, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
					else:
						print('Processo Finalizado')

				if ncarteira == 5:
					print(
						'Criptomoeda: Theter (USDT), Quantidade: {}, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%'.format(
							saldo5, al5))
					print('')
					transfer = str(input('Deseja movimentar o circuito interno? '))
					print('')
					if transfer == 'não'.upper().lower() or transfer == 'nao'.upper().lower():
						print('')
					if transfer == 'sim'.upper().lower():
						print('')
						valor = float(input('Qualifique a tranferencia: '))
						print('')
						destino = leiaNcarteira('Qual a carteira de destino? (digite o número da carteira): ')
						# taxa = random.choice(
						# 	[0.001562, 0.002604, 0.002322, 0.002298, 0.002512, 0.002233, 0.001666, 0.001732, 0.001878,
						# 	 0.001904,
						# 	 0.002043, 0.002179, 0.002417, 0.002471, 0.002438, 0.001763, 0.001589, 0.001622, 0.001927,
						# 	 0.001856])
						# taxa2 = taxa * 10
						taxaUsdt = random.choice(
							[1.77289, 3.86556, 4.29802, 11.11901, 7.01632, 9.144054, 6.00988, 1.39284, 8.85470,
							 13.39448])

						taxaUsdt2 = random.choice(
							[25.41013, 36.41005, 63.033228, 32.902384, 40.40544, 55.07471, 61.91639, 27.45124, 49.73885,
							 67.16942])
						print('')
						if destino == 1 or destino == 2 or destino == 3 or destino == 4:
							print(
								'''	
                            ERC-20 Transaction: Tax pending.
                            Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                            Conversão será feita automaticamente via protocolo crosschain (1 INCH NETWORK).
                            ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format(na1, na2, na3, la1, na4, na5, na6, na7, na8, la2,
																   la3))
							time.sleep(3)
							confirmar = leiaCriptografia('''
							O valor de taxa de é {} USDT.
							Digite sim, para confirmar:'''.format(taxaUsdt2))
							if confirmar == 'sim'.upper().lower().strip():
								saldo5 = saldo5 - valor
								if destino == 1:
									saldo1 = saldo1 + (valor / 1794) - taxaUsdt2
									real1 = float(saldo1 * 9097.38)
									al1 = round((100 * real1) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 2:
									saldo2 = saldo2 + (valor / 1794) - taxaUsdt2
									real2 = float(saldo2 * 9097.38)
									al2 = round((100 * real2) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 3:
									saldo3 = saldo3 + valor - taxaUsdt
									real3 = float(saldo3 * 9097.38)
									al3 = round((100 * real3) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 4:
									saldo4 = saldo4 + valor - taxaUsdt
									real4 = float(saldo4 * 9097.38)
									al4 = round((100 * real4) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 5:
									saldo5 = saldo5 + valor - taxaUsdt
									real5 = (saldo5 * 5.07)
									saldo5 = round(saldo5, 5)
									al5 = round((100 * real5) / total, 3)
								if destino == 6:
									saldo6 = saldo6 + valor - taxaUsdt
									real6 = (saldo6 * 5.07)
									saldo6 = round(saldo6, 5)
									al6 = round((100 * real6) / total, 3)
								for c in range(1, ta1):
									print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta2):
									print('Confirmação de rede pendende (1/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta3):
									print('Confirmação de rede pendende (2/3): PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								print('Confirmação de rede pendende (3/3): PROTOCOLO -',
									  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))

								if destino == 1:
									print(
										'O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo1, al1, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 2:
									print(
										'O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo2, al2, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 3:
									print(
										'O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo3, al3, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 4:
									print(
										'O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo4, al4, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 5:
									print(
										'O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo5, al5, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 6:
									print(
										'O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo6, al6, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))

						else:

							print(
								'''	
                            ERC-20 Transaction: Tax pending.
                            Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                            ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format(na1, na2, na3, la1, na4, na5, na6, na7, na8, la2,
																   la3))
							time.sleep(3)
							confirmar = leiaCriptografia('''
							O valor de taxa de é {} USDT.
							Digite sim, para confirmar:'''.format(taxaUsdt))
							if confirmar == 'sim'.upper().lower().strip():
									saldo5 = saldo5 - valor
									if destino == 1:
										saldo1 = saldo1 + (valor/1794) - taxaUsdt2
										real1 = float(saldo1 * 9097.38)
										al1 = round((100 * real1) / total, 3)
										for c in range(1, 60):
											print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
												  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
											time.sleep(1.5)
									if destino == 2:
										saldo2 = saldo2 + (valor/1794) - taxaUsdt2
										real2 = float(saldo2 * 9097.38)
										al2 = round((100 * real2) / total, 3)
										for c in range(1, 60):
											print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
												  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
											time.sleep(1.5)
									if destino == 3:
										saldo3 = saldo3 + valor - taxaUsdt
										real3 = float(saldo3 * 9097.38)
										al3 = round((100 * real3) / total, 3)
										for c in range(1, 60):
											print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
												  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
											time.sleep(1.5)
									if destino == 4:
										saldo4 = saldo4 + valor - taxaUsdt
										real4 = float(saldo4 * 9097.38)
										al4 = round((100 * real4) / total, 3)
										for c in range(1, 60):
											print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
												  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
											time.sleep(1.5)
									if destino == 5:
										saldo5 = saldo5 + valor - taxaUsdt
										real5 = (saldo5 * 5.07)
										saldo5 = round(saldo5, 5)
										al5 = round((100 * real5) / total, 3)
									if destino == 6:
										saldo6 = saldo6 + valor - taxaUsdt
										real6 = (saldo6 * 5.07)
										saldo6 = round(saldo6, 5)
										al6 = round((100 * real6) / total, 3)
									for c in range(1, ta1):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
									for c in range(1, ta2):
										print('Confirmação de rede pendende (1/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
									for c in range(1, ta3):
										print('Confirmação de rede pendende (2/3): PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
									print('Confirmação de rede pendende (3/3): PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))

									if destino == 1:
										print(
											'O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo1, al1, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 2:
										print(
											'O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo2, al2, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 3:
										print(
											'O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo3, al3, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 4:
										print(
											'O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo4, al4, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 5:
										print(
											'O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo5, al5, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 6:
										print(
											'O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo6, al6, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
					else:
						print('Processo Finalizado')

				if ncarteira == 6:
					print(
						'Criptomoeda: Theter (USDT), Quantidade: {}, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%'.format(
							saldo6, al6))
					print('')
					transfer = str(input('Deseja movimentar o circuito interno? '))
					print('')
					if transfer == 'não'.upper().lower() or transfer == 'nao'.upper().lower():
						print('')
					if transfer == 'sim'.upper().lower():
						print('')
						valor = float(input('Qualifique a tranferencia: '))
						print('')
						destino = leiaNcarteira('Qual a carteira de destino? (digite o número da carteira): ')
						# taxa = random.choice(
						# 	[0.001562, 0.002604, 0.002322, 0.002298, 0.002512, 0.002233, 0.001666, 0.001732, 0.001878,
						# 	 0.001904,
						# 	 0.002043, 0.002179, 0.002417, 0.002471, 0.002438, 0.001763, 0.001589, 0.001622, 0.001927,
						# 	 0.001856])
						# taxa2 = taxa * 10
						taxaUsdt = random.choice(
							[1.77289, 3.86556, 4.29802, 11.11901, 7.01632, 9.144054, 6.00988, 1.39284, 8.85470,
							 13.39448])

						taxaUsdt2 = random.choice(
							[25.41013, 36.41005, 63.033228, 32.902384, 40.40544, 55.07471, 61.91639, 27.45124, 49.73885,
							 67.16942])
						print('')
						if destino == 1 or destino == 2 or destino == 3 or destino == 4:
								print(
									'''	
                                ERC-20 Transaction: Tax pending.
                                Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                                Conversão será feita automaticamente via protocolo crosschain (1 INCH NETWORK).
                                ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format(na1, na2, na3, la1, na4, na5, na6, na7, na8,la2, la3))
								time.sleep(3)
								confirmar = leiaCriptografia('''
								O valor de taxa de é {} USDT.
								Digite sim, para confirmar:'''.format(taxaUsdt2))

								if confirmar == 'sim'.upper().lower().strip():
									saldo6 = saldo6 - valor
									if destino == 1:
										saldo1 = saldo1 + (valor / 1794) - taxaUsdt2
										real1 = float(saldo1 * 9097.38)
										al1 = round((100 * real1) / total, 3)
										for c in range(1, 60):
											print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
												  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
											time.sleep(1.5)
									if destino == 2:
										saldo2 = saldo2 + (valor / 1794) - taxaUsdt2
										real2 = float(saldo2 * 9097.38)
										al2 = round((100 * real2) / total, 3)
										for c in range(1, 60):
											print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
												  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
											time.sleep(1.5)
									if destino == 3:
										saldo3 = saldo3 + valor - taxaUsdt
										real3 = float(saldo3 * 9097.38)
										al3 = round((100 * real3) / total, 3)
										for c in range(1, 60):
											print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
												  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
											time.sleep(1.5)
									if destino == 4:
										saldo4 = saldo4 + valor - taxaUsdt
										real4 = float(saldo4 * 9097.38)
										al4 = round((100 * real4) / total, 3)
										for c in range(1, 60):
											print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
												  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
											time.sleep(1.5)
									if destino == 5:
										saldo5 = saldo5 + valor - taxaUsdt
										real5 = (saldo5 * 5.07)
										saldo5 = round(saldo5, 5)
										al5 = round((100 * real5) / total, 3)
									if destino == 6:
										saldo6 = saldo6 + valor - taxaUsdt
										real6 = (saldo6 * 5.07)
										saldo6 = round(saldo6, 5)
										al6 = round((100 * real6) / total, 3)

									for c in range(1, ta1):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
									for c in range(1, ta2):
										print('Confirmação de rede pendende (1/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
									for c in range(1, ta3):
										print('Confirmação de rede pendende (2/3): PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
									print('Confirmação de rede pendende (3/3): PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))

									if destino == 1:
										print(
											'O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo1, al1, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 2:
										print(
											'O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo2, al2, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 3:
										print(
											'O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo3, al3, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 4:
										print(
											'O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo4, al4, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 5:
										print(
											'O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo5, al5, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
									if destino == 6:
										print(
											'O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
												saldo6, al6, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
						else:
							print(
								'''	
                            ERC-20 Transaction: Tax pending.
                            Aviso: Taxa de gás inclui crosschain dinâmico proporcional ao gás.
                            ID: {}{}{}{}{}{}{}{}{}-{}{} '''.format(na1, na2, na3, la1, na4, na5, na6, na7, na8, la2,
																   la3))
							time.sleep(3)
							confirmar = leiaCriptografia('''
							O valor de taxa de é {} USDT.
							Digite sim, para confirmar:'''.format(taxaUsdt))
							if confirmar == 'sim'.upper().lower().strip():
								saldo6 = saldo6 - valor
								if destino == 1:
									saldo1 = saldo1 + (valor / 1794) - taxaUsdt2
									real1 = float(saldo1 * 9097.38)
									al1 = round((100 * real1) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 2:
									saldo2 = saldo2 + (valor / 1794) - taxaUsdt2
									real2 = float(saldo2 * 9097.38)
									al2 = round((100 * real2) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 3:
									saldo3 = saldo3 + valor - taxaUsdt
									real3 = float(saldo3 * 9097.38)
									al3 = round((100 * real3) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 4:
									saldo4 = saldo4 + valor - taxaUsdt
									real4 = float(saldo4 * 9097.38)
									al4 = round((100 * real4) / total, 3)
									for c in range(1, 60):
										print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
											  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
										time.sleep(1.5)
								if destino == 5:
									saldo5 = saldo5 + valor - taxaUsdt
									real5 = (saldo5 * 5.07)
									saldo5 = round(saldo5, 5)
									al5 = round((100 * real5) / total, 3)
								if destino == 6:
									saldo6 = saldo6 + valor - taxaUsdt
									real6 = (saldo6 * 5.07)
									saldo6 = round(saldo6, 5)
									al6 = round((100 * real6) / total, 3)

								for c in range(1, ta1):
									print('Confirmação de rede pendende (0/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta2):
									print('Confirmação de rede pendende (1/3):  PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								for c in range(1, ta3):
									print('Confirmação de rede pendende (2/3): PROTOCOLO -',
										  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))
									time.sleep(1.5)
								print('Confirmação de rede pendende (3/3): PROTOCOLO -',
									  ''.join(random.SystemRandom().choices(string.ascii_letters, k=7)))

								if destino == 1:
									print(
										'O saldo da carteira R1 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo1, al1, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 2:
									print(
										'O saldo da carteira R2 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo2, al2, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 3:
									print(
										'O saldo da carteira R3 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo3, al3, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 4:
									print(
										'O saldo da carteira R4 é {} em PAXG, Status: Chaves ativas sob gerenciamento, Tipo: Reserva, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo4, al4, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 5:
									print(
										'O saldo da carteira I1 é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Intermediaria, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo5, al5, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
								if destino == 6:
									print(
										'O saldo da carteira FTX é {} em USDT, Status: Chaves ativas sob gerenciamento, Tipo: Corretora, AL: {}%, ID: {}{}{}{}{}{}{}{}{}-{}{}'.format(
											saldo6, al6, na1, na2, na3, la1, na4, na5, na6, na7, na8, la2, la3))
					else:
						print('Processo Finalizado')


	# informações criptografia
	if opcao == 5:
			print("")
			print("Informações de criptografia: ")
			print("")
			print("")
			print("")
			print("")
			print("O sistema possui capacidade de encriptografamento de padrão Advanced Encryption Standard (AES) ")
			print("")
			print('''
		Possuimos uma criptografia assimétrica que possui HASHs de segurança, 
		e uma cifra de fluxo (stream cipher), pois é um dos melhores e mais 
		seguros protocolos de criptografia e é utilizado em incontáveis aplicações.''')
			print("")
			print("")
			print('''
		A tecnologia utilizada na criptografia são protocolos Secure Socket Layer (SSL),
		hoje conhecido como Transport Layer Security (TLS), 
		Derivada do DES, a criptografia 3DES (ou Triplo DES).''')
			print("")
			print("")
			criptousuario = str(input("Digite uma mensagem que gostaria que fosse encriptografada: "))
			print("")
			print("Sua mensagem encriptografada ficou: ")
			print("")
			print(criptografia(criptousuario))
			print("")


	# informações protocolo de segurança
	if opcao == 6:
		print("")
		print("Informações Protodolo de segurança: ")
		print("")
		print("")
		print("")
		print("")
		print(
			'''O sistema possui segurança de encriptografamento de padrão Advanced Encryption Standard (AES), 
ID de rede classe dhcp, IPv6 permitidas para o adaptador com conexão segura e protocolos de segurança WEP e WPA''')
		print("")
		print('''
		Wired Equivalent Privacy  (Privacidade equivalente aos fios) WEP é um sistema de criptografia adotado pelo padrão IEEE 802.11. 
		Ele utiliza uma senha compartilhada para criptografar os dados e funciona de forma estática. 
		Além de fornecer apenas um controle de acesso e de privacidade de dados na rede sem fio.
				''')
		print("")
		print("")
		print('''
		Wi-fi Protected Access  (Wi-Fi de acesso protegido) Também é conhecido como TKIP (Temporal Key Integrity Protocol). 
		O recurso surgiu em 2003 para aumentar a segurança do protocolo WEP. O WPA2 é considerada a versão final o WPA. 
		A principal diferença entre o WPA e o WPA2 é a forma com a qual ele criptografa os dados. 
		Enquanto o WPA utiliza o TKIP como algoritmo de criptografia, o WPA2 utiliza o algoritmo AES (Advanced Encryption Standard). 
		O algoritmo AES é consideravelmente mais pesado que o TKIP. 
		Por conta disso, as placas mais antigas não suportam o WPA2, nem com um firmware atualizado.
		O AES é o padrão de criptografia utilizado''')
		print("")
		print("")


	# if opcao == 3:
	# 	print("")
	# 	print("")
	# 	print("")
	# 	print("")
	# 	print("A aplicação está consultando as cotações atuais...")
	# 	print("")
	#
	# 	# As coordenadas X e Y variam conforma a resolução
	# 	# Para descobrir as coordenadas do cursor do mouse usar:
	# 	pyautogui.PAUSE = 2
	#
	# 	pyautogui.click(button="right", x=4052, y=1061)
	# 	pyautogui.click(x=4067, y=868)
	# 	time.sleep(5)
	# 	print(pyautogui.position())

