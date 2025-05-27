from avaliacao import Avaliacao #Importanto outro arquivo local (igual o import de php)

#########################----------------------Boas Praticas----------------------#########################
'''
Como indicar que algo é uma constante em Python? Escrevendo o nome em letras maiúsculas. (o valor nao deve ser alterado, nao signifca que nao possa,
mas sim que nao devemos mexer naquela variavel) 
'''

###########Docstring
'''
-usada para facilitar/documentar módulo, função, classe ou método em Python através de uma descricao.
Ela é colocada como o primeiro item de definição e pode ser acessada usando a função help() ou apenas deixando o mouse acima do elemento.

-deve descrever propósito, parâmetros, tipo de retorno e exceções levantadas pela função

-facilitar a leitura, manutenção e compartilhamento do código com outras pessoas desenvolvedoras.
'''
#Exemplo..
def media333(lista: list=[0]) -> float:
  ''' Função para calcular a média de notas passadas por uma lista

  lista: list, default [0]
    Lista com as notas para calcular a média
  return = calculo: float
    Média calculada
  '''
  calculo = sum(lista) / len(lista)
  return calculo
media333()
help(media333)

###########Funcao Lambda (funcao anonima)
#não precisa ser definida(não possui um nome)
#Pode ser descrita uma única linha.
# Funcao de soma em formato lambda
nota = float(input("Digite a nota do(a) estudante: "))
soma_05_na_nota = lambda x: x + 0.5 #X e o paramtro aqui
soma_05_na_nota(nota) # nota sera passado como parametro ("x")



#########################----------------------Codigos Gerais----------------------#########################

#Criando variavel e atribuindo valor atraves de informacao inserido pelo usuario
variavel_informacao_inserida = int(input('Insira um valor numerico: ')) #int() transforma string em number
#input retorna o valor como uma string
#se verificarmos o tipo com   print(TYPE(opcao_escolhida)) retornara <class 'str'> = string, se verificarmos dps do int ela estara como number

#o sep (separador) que é de espacos (' ') agora será de pular linhas (/n) 
print('A','L','U','R','A',sep='\n')

#f strings --> igual a template string do javascript
f_strings = print(f'texto seguido de codigo ou variavel tudo em um mesmo lugar --> {variavel_informacao_inserida}')

#Para colocar x espacamentos, podendo ser ljust(x) espacamentos a esquerda, rjust(x) espacamentos a direita
texto_para_testes = 'teste'
print(f'Informacoes com espacamento - {texto_para_testes.rjust(10)} | {texto_para_testes.ljust(10)}')

#Para inverter o valor de algo basta colocar "not" antes ex:
print(not True) 




#########################----------------------Bibliotecas----------------------#########################
#Conj de modulos de funcoes prontas, realizam tarefas sem que seja necessário trazer todos os pacotes e dados para o projeto.

import os #carregando uma biblioteca do pyton
os.system('cls') #utilizando uma funcao da biblioteca os do python
#cls para windows limpa o terminal

#Para instalar ou atualizar uma biblioteca, podemos utilizar o pip
############### INSTALANDO bibliootecas (atraves do terminal) ############################## (obs no collab precisa de um ! antes do pip)
#pip install matplotlib  --> inserir isso no terminal
#pip install matplotlib==3.6.2 --> para instalar um versao especifica

#import numpy as np   os as e apenas para "apelidar" e podermos usar o numpy atraves desse apelido
# dentro de matplotlib, tem modulos, submodulos e outras bibliotecas(como o numpy), isso n transforma o numpy em modlo ou submodulo, ele segue senod uma biblioteca

##--Importando uma função específica de uma biblioteca--##
#from random import randrange, sample, outro, metodo --> para importar um ou mais módulos/duncoes

from random import choice #----> economiza memoria importar apenas oq precisamos e nao a biblioteca toda

#help(choice) # Retorna uma explicacao de como funciona e do que precisa esse modulo

#plt.bar(x = estudantes, height = notas) #para criar graficos a partir dos dados, utilizaremos metodos da biblioteca que importamos
#metodo bar()
#plt.show() --> para mostrar o grafico criado

#########################----------------------strings----------------------#########################
string = 'texto teste para strings'
print(len('esse texto tem 70 caracteres, considerando tambem os espacos em branco'))
#float(input('Insira um numero: '))   -->    float converte string em numero(aceita casas decimais)
#int(input('Insira um numero: '))     -->    int tambem converte, porem nao aceita casas decimais, apenas numero inteiros
string.title() #Torna a primeira letra das palavras maiusculas
string.upper() #Torna todas letras maiusculas desses atributos
string.lower() #Torna todas letras minusculas desses atributos
string.strip() #retira todos os espacos " " entre as palavras
string.replace('e',chr(64)) #trocando 'e' por 'chr(64)' --> e a forma de exibir caracteres especias em python, se pesquisar tem os valroes na net
string.split() #divide uma string em uma lista de varias strings com base em um delimitador específico, nesse caso é o () espaco em branco, mas podemos escolher o elemento que quisermos


vogais = 'aeiou' # string contendo todos os dados
# verifica se tem o elemento vogais em(in) string
if string in vogais:
    print('A letra é uma vogal.')
else:
    print('A letra é uma consoante.')


#########################----------------------Numbers----------------------#########################
#str()  Converte um valor para string




#########################----------------------Funcoes----------------------#########################
#funcoes built-in: Funções internas, elas ja vem definidas/integradas na linguagem de programação
#elas estão sempre dispostas para uso, ex de funcoes: print(), len(), type(), int()...

#Criando uma funcao 
def main(): print('teste')
if __name__ == '__main__':
    main()

"""
`__name__:` é uma variável especial que recebe o nome do módulo/arquivo (nesse caso é python.py).
Quando o codigo esta dentro do módulo e é executado, __name__ é definido como '__main__'.
Se o módulo for importado, ou seja, o codigo feito aqui, for importado e executado em outro arquivo/modulo,
este __name__ acima, será definido como o nome do módulo onde ele foi executado (python.py), e nao mais como main,
assim evitando a execucao desse codigo de verificacao if.
"""

#texto: list, int... o que vier apos os dois pontos serve apenas para indicar o que o parametro que esta ali
# serve apenas para melhor compreensao
def divide_colunas(lista_1: list, lista_2: list) -> list:
    print() #Ja a -> serve tbm apenas para indicar, porem dessa vez para indicar o tipo de informacao que deve retornar



#########################----------------------Dicionario----------------------#########################
#Dicionário {chave: valor}
dicionario_basico = {'nome': 'Amanda', 'idade': 19, 'cidade': 'São Luís'}
#para atribuir mais de um valor a cada chave devemos colocar os valores em listas ou tuplas --> ex: dicionario_basico = {'nome': [19, 'São Luís']}

#verificando se um elemento esta presente no dicionario
if 'nome' in dicionario_basico:
    print("A chave 'nome' existe no dicionário.")
else:
    print("A chave 'nome' não existe no dicionário.")
    
# Adicionando Profissão já com um valor
dicionario_basico['profissao'] = 'Engenheiro'

lista_de_dicionarios = [{'nome':'Praça', 'categoria':'Japonesa', 'ativo':False}, # Cada um desses e um dicionario ou seja aqui tem 3 dicionarios dentro de uma lista
                {'nome':'Pizza Suprema', 'categoria':'Pizza', 'ativo':True},
                {'nome':'Cantina', 'categoria':'Italiano', 'ativo':False}]
#Forma de atualizar um dicionario -->  restaurantes.update({'nome': 'ta ai'})
# Alterando o valor da chave 'nome' do primeiro dicionário
lista_de_dicionarios[0]['nome'] = 'Praça Central'

# Alterando o valor da chave 'ativo' do segundo dicionário
lista_de_dicionarios[1]['ativo'] = False
# Alterando o título da chave 'nome' para 'titulo' em todos os dicionários
for i in lista_de_dicionarios:
    i['titulo'] = i.pop('nome') #.pop usado para remover um item de um dicionário (ou lista) e retornar o valor associado a esse item

#para atualizar um valor
pessoa.update({'idade': 18})

# Adicionando um novo dicionario
lista_de_dicionarios.append({'nome': 'Café da Esquina', 'categoria': 'Café', 'ativo': True})



#########################----------------------LISTAS E TUPLAS------------------#######################
#lista = [1,’olá mundo’,True,9.7] --> sao mutaveis, podemos adc,excluir, alterar elementos
#tupla = (1,’olá mundo’,True,9.7) --> sao imutaveis
restaurantes = []
restaurantes.append(input('Digite o nome do restaurante que deseja cadastrar: ')) 
#Adicionando elemento dentro de uma lista, OBS somente a lista aceita isso pois ela e mutavel, a tupla nao aceita insercao de dados dessa maneira

#para verificar se um elemento esta presente em uma lista
'''lista_exemplo = [1, 2, 3, 4, 5, 6]
presente_na_lista = 4
nao_presente_na_lista = 9
if presente_na_lista in lista_exemplo:
    print('presente')
if nao_presente_na_lista not in lista_exemplo:
    print('nao esta presente')'''
    
notas333 = {'19 Trimestre': 8.5, '2° Trimestre': 9.5, '3º trimestre': 7}
soma = 0
for nota in notas333.values(): #values: transforma os valores em uma lista iterável (que mode ser modificada, mexida)
    soma += nota

# Não conseguimos aplicar o lambda em lista direto, é necessário utilizarmos junto a ela a função map
notas222 = [6.0, 7.0, 9.0, 5.5, 8.0]
qualitativo2 = 0.5
notas_atualizadas = map(lambda x: x + qualitativo2, notas222) #se tentar printar aqui ira rolar <map at 0x7f57f633b5d0> objeto do tipo mapa
#que mapeou os valores. Para conseguimos visualizá-lo, precisamos usar list que transforma o objeto mapa em uma lista.
notas_atualizadas = list(notas_atualizadas)
print(notas_atualizadas)


'''
Sobre Tuplas
As tuplas são estruturas de dados imutáveis, são aplicadas para agrupar dados que não devem ser
modificados. Ou seja, não é possível adicionar, alterar ou remover seus elementos depois de criadas.
apena podemos utiliza-los para resolver outras coisas mas sem nunca mexer neles dentro da tupla em si
'''
#Tuplas sao semelhantes a listas a diferenca e que usamos () ao inves de []


#funcao filter --> recebe um elemento/lista/..., nesse caso e "frase" e aplica uma funcao para comprar ha uma regra/filtro e ira retornar true ou false, como nesse caso e uma lista ela ira iterra sobre cada elemento
frase = 'apenas palavras de cinco ou mais elementos serao aceitos'
lista_frases = list(filter(lambda iteravel: len(iteravel)  >= 5, frase))
#A regra aqui é qie apenas elementos de 5 ou mais caracteres serao aceitos

#funcao set --> pega apenas valores unicos de uma lista/tupla com valores repetidos
estados_unicos = list(set(estados)) #comparando o elementod e uma lista
estados_unicos = list(set([tupla[0] for tupla in funcionarios])) #comparando cum elemento de uma lista de lista

#########################----------------------Listas de Listas------------------#######################
notas_turma = ['João', 8.0, 9.0, 10.0, 'Maria', 9.0, 7.0, 6.0, 'José', 3.4, 7.0, 7.0,
               'Cláudia', 5.5, 6.6, 8.0, 'Ana', 6.0, 10.0, 9.5]

nomes = []
notas_juntas = []

#Dividindo a lista de listas em duas listas distintas
for i in range(len(notas_turma)): 
  if i % 4 ==0: #eesa expressao verifica se o "i"(indice) em questao e multiplo de 4 --> 0, 4, 8, 12, 16, ...
      nomes.append(notas_turma[i]) #nomes = [] --> ['João', 'Maria', 'José', 'Cláudia', 'Ana']
  else: 
      notas_juntas.append(notas_turma[i])#notas_juntas = [] --> [8.0, 9.0, 10.0, 9.0, 7.0, 6.0, 3.4, 7.0, 7.0, 5.5, 6.6, 8.0, 6.0, 10.0, 9.5]

notas = []

#colocamos como parametro onde comeca (0), onde termina (len(notas_juntas)), e o valor de incremento(3) quantidade de iontervalo que ele pula
for i in range(0, len(notas_juntas), 3): 
  notas.append([notas_juntas[i], notas_juntas[i+1], notas_juntas[i+2]]) # adc listas em uma lista
#a logica aqui e a seguinte: ele adiciona os elementos de indice 0, 0+1, 0+2 apos ele pula as 3 casas
#de incremento padrao dele, entao na proxima sequencia ele ira adiocionar os elementos de indice 3, 3+1, 3+2 e assim sucessivamente

#vira uma lista de listas
#notas[] --> [[8.0, 9.0, 10.0], [9.0, 7.0, 6.0], [3.4, 7.0, 7.0], [5.5, 6.6, 8.0], [6.0, 10.0, 9.5]]

estudantes = ["João", "Maria", "José", "Cláudia", "Ana"]
codigo_estudantes = []
from random import randint #pegando a funcao randint da biblioteca random
def gera_codigo():
  return str(randint(0,999)) #randint retorna um valor aleatorio INTEIRO aleatorio dentro do intervalo estabelecido
    #ambos os numeros sao inclusivos

#vamos fazer um loop de repeticao de 0 ate a quantidade de estudantes dentro da lista
for i in range(len(estudantes)): 
  codigo_estudantes.append((estudantes[i], estudantes[i][0] + gera_codigo()))
  #iremos colocar dentro da lista codigo_estudantes a seguinte tupla
  #estudantes[i] --> nome relativo ao indice "i",   estudantes[i][0] --> elemento/letra de indice[0] do nome de indice "i", por fim soma essa letra com o codigo aleatorio gerado pela funcao
estudantes
print(codigo_estudantes) #--> [('João', 'J482'),('Maria', 'M576'), ('José', 'J929'), ('Cláudia', 'C736'), ('Ana', 'A505')]


#########################----------------------Objetos/classes------------------#######################
'''Uma classe é um molde para criar objetos que compartilham as mesmas propriedades e comportamentos.'''
#OBS: Criar classes com iniciais maiusculas e uma boa pratica
#funcoes criadas dentro de classses sao chamadas de métodos de instancia, pódendo ser metodos normais, estasticos...

#Exmeplo basico de classe
class Exemplo:
    nome = ''
    ativo = False

#CRIANDO OBJETOS BASEADO NA CLASSE ACIMA
restaurante_praca = Exemplo()
restaurante_praca.nome = 'Praça'

print(dir(restaurante_praca)) #dir --> retorna informacoes sobre o elemento inserido, funcoes, metodos, atribustos...
print(vars(restaurante_praca)) #vars --> retorna em formato de dicionario o objeto criado



#Exemplo mais completo e complexo
class Restaurante: #Cria uma nova instância da classe Restaurante chamada restaurante_praca...
#...Uma instância é um objeto específico criado a partir de uma classe.

    def __init__(self, nome, categoria): #O método __init__ é o construtor da classe, 
        #é chamado toda vez que um novo objeto da classe Restaurante é criado
        self._nome = nome #o self, e como um this do javascript ou "i" de loopings, podemos colocar o nome que quisermos, ele servira como uma referencia
        self._categoria = categoria
        self._ativo = False
        #O "_" Uma convencao para indicar que nao devemos mexer naquele atributo, atributo privado
        self._avaliacao = []
        Restaurante.restaurantes.append(self)
        #Toda vez que criarmos um objeto/instancia ele sera inserido dentro da lista/array 'restaurantes'
    
    #__str__ esse metodo define como sera a saida do bjeto em formatod e string
    def __str__(self):
        return f'{self._nome} | {self._categoria}' #nesse caso sera mostrando os respectivos valores seprados por "|"
    
    
    #cls: Parâmetro que representa a própria classe. Semelhante ao self em métodos de instância,
    #porem self se refere a uma instância específica da classe, cls se refere à classe como um todo.
    
    @classmethod #Se for um metodo referenciado com alguma classe, e nao com um objeto (__init__, __str__,
    #alternar_estado), devemos utilizar esse @classmethod, e uma convencao para quem ler o codigo saber
    #diferenciar ao que um metodo pertence, instancia/objeto ou a classe
    def listar_restaurantes(cls):
        print(f'{'Nome do restaurante'.ljust(25)} | {'Categoria'.ljust(25)} | {'Avaliação'.ljust(25)} |{'Status'}')
        for restaurante in cls.restaurantes:
            print(f'{restaurante._nome.ljust(25)} | {restaurante._categoria.ljust(25)} | {str(restaurante.media_avaliacoes).ljust(25)} |{restaurante.ativo}')
   #Para acessarmos um metodo de classe e como acessar qualquer outro metodo basta: colocarmos o elemento.metodo ex --> Restaurante.listar_restaurantes()
   
    '''
        métodos que pertencem à classe, mas não têm acesso à instância (self) ou à classe (cls).
        Eles são usados quando a lógica do método não depende do estado da instância ou da classe,
        mas ainda está logicamente relacionado à classe.

        @staticmethod 
        def verificar_disponibilidade(ano):
            livros_disponiveis = [livro for livro in Livro.livros if livro.ano_publicacao == ano and livro.disponivel]
            return livros_disponiveis
    '''
    
    @property
    def ativo(self):
        return '⌧' if self._ativo else '☐'
    
    def alternar_estado(self):
        self._ativo = not self._ativo

    def receber_avaliacao(self, cliente, nota):
        avaliacao = Avaliacao(cliente, nota)
        self._avaliacao.append(avaliacao)

    @property #para deixarmos esse metodo (media_avaliacoes) disponivel para utilizar-lo fora da class o transformamos em property
    def media_avaliacoes(self):
        if not self._avaliacao:  #se o restaureante em questao (self) nao ter nenhuma avaliacao return 0
            return 0              #para cada(for)avaliacao na(in) nossa lista self._avaliacao
        soma_das_notas = sum(avaliacao._nota for avaliacao in self._avaliacao)
        quantidade_de_notas = len(self._avaliacao) #len (lenght) retorna a quantidade de avaliacoes que temos na lista(_avaliacao)
        media = round(soma_das_notas / quantidade_de_notas, 1) #round serve para arredonar para numero de casas decimais definidos
        #       round(valor que sera arredondado, quantidade de casas decimais permitidas)
        return media





#O método __init__ é o construtor da classe. Ele é chamado toda vez que um novo objeto da classe Restaurante é criado. Vamos detalhar o que acontece dentro dele:


#Criando um novo objeto, baseado na classe Restaurante
restaurante_praca = Restaurante()
restaurante_praca.nome = 'Praça' #Define um valor para o atributo "nome"
restaurante_praca.categoria = 'Gourmet'



#########################-------------------Loopings / laços de repetição-------------------#########################
for restaurante in restaurantes: #para cada "i" em "lista/tupla": Faca codigo...
    print(f'.{restaurante}')     #para cada "restaurante" em "restaurantes": Faca codigo print('bla')

"""
while para quando nao sabemos quantas repeticoes iremos precisar

FOR normal, para quando sabemos quantas repeticoes teremos
for i in range(3, 6) --> assim ira pegar todos os numeros de 3 a 5,   sempre o primeiro numero(inclusivo) e os segundo(exclusivo)
for i in range(3):  # Número máximo de tentativas (3)
    numero = int(input("Digite um número positivo: "))
    if numero > 0:
        break Encerra o ciclo de repeticao

print("Você digitou:", numero)
"""

#Switch case do python / match case
"""
variavel_informacao_inserida = int(input('Escolha uma opção: '))
match variavel_informacao_inserida:
    case 1:
        print('Adicionar restaurante')
    case 2:
        print('Listar restaurantes')
    case 3:
        print('Ativar restaurante')
    case 4:
        print('Finalizar app')
    case _:
        print('Opção inválida!')
"""




#########################----------------------Try/Except----------------------#########################
'''
Try: codigo --> ira executar o codigo normalmente caso de um erro ele executa o codigo do except
Excpet: codigo --> ira executar caso ocorra algum erro no codigo do try acima dele
'''
try:
    opcao_escolhida = int(input('Escolha uma opção: '))
    # opcao_escolhida = int(opcao_escolhida)
        
    if opcao_escolhida == 1: 
        print('1')
    elif opcao_escolhida == 2: 
        print('2')
    else:
        print('3')
        #"as" serve como no SQL, para "apelidar/apreviar" o nome de algo, logo nesse caso e = Exception
except Exception as e: #Exception ira abordar todos os erros (KeyError, ValueError, NameError, IndexError...)
    print(type(e), f"Erro: {e}") ##...<class 'KeyError'> Erro: 'Mirla'

##----------------Tipos de Exceções----------------##
#1SyntaxError (Erro de escrita): uma seta aponta para a parte do código que gerou o erro

''' Exemplo
  File "<ipython-input-16-2db3afa07d68>", line 1
    print(10/2
              ^
SyntaxError: unexpected EOF while parsing
'''

#2NameError: quando tentamos utilizar algum elemento que não está presente em nosso código.
'''raiz = sqrt(100)
---------------------------------------------------------------------------
NameError                                 Traceback (most recent call last)
<ipython-input-17-2e14e900fb9f> in <module>
----> 1 raiz = sqrt(100)

NameError: name 'sqrt' is not defined
'''


#3IndexError: quando tentamos indexar alguma estrutura de dados como lista, tupla ou até string além de seus limites.
'''
lista = [1, 2, 3]
lista[4]
---------------------------------------------------------------------------
IndexError                                Traceback (most recent call last)
<ipython-input-18-f5fe6d922eea> in <module>
      1 lista = [1, 2, 3]
----> 2 lista[4]

IndexError: list index out of range
'''


#4TypeError: quando um operador ou função são aplicados a um objeto cujo tipo é inapropriado.
'''
"1" + 1
---------------------------------------------------------------------------
TypeError                                 Traceback (most recent call last)
<ipython-input-20-ec358fc6499a> in <module>
----> 1 "1" + 1

TypeError: can only concatenate str (not "int") to str
'''


#5KeyError: quando tentamos acessar uma chave que não está no dicionário presente em nosso código.
'''
estados = {'Bahia': 1, 'São Paulo': 2, 'Goiás': 3}
estados["Amazonas"]

---------------------------------------------------------------------------
KeyError                                  Traceback (most recent call last)
<ipython-input-22-45729db26889> in <module>
      1 estados = {'Bahia': 1, 'São Paulo': 2, 'Goiás': 3}
----> 2 estados["Amazonas"]

KeyError: 'Amazonas'

'''


####-------------------------------finally-----------------------------####
'''
finally: codigo --> codigo que sera executado independente se houver erro ou nao ex:
'''
try:
    nome = input("Digite o nome do(a) estudante: ")
    resultado = notas[nome]
except KeyError:
    print("Estudante não matriculado(a) na turma")
else:
    print(resultado)
finally:
    print("A consulta foi encerrada!")
    
    
####-------------------------------Raise-----------------------------####
'''
O raise é usado para lançar uma exceção propositalmente. Isso interrompe a execução normal
do código e direciona o fluxo geralmente para o bloco except, onde o erro é tratado.
'''
''' Função para calcular a média de notas passadas por uma lista
lista: list, default[0]
lista: list, default[0]
    Lista com as notas para calcular a média
    return calculo: float
    Média calculada
'''
def media(lista: list=[0]) -> float:
    calculo = sum(lista) / len(lista)

    if len(lista) > 4:
        raise ValueError("A lista não pode possuir mais de 4 notas.")
    
    return calculo

try:
    notas = [6, 7, 8, "9"]
    resultado = media(notas)
except TypeError:
    print("Não foi possível calcular a média do(a) estudante. Só são aceitos valores numéricos!")
except ValueError as e: #lembra que "as" e apenas para apelidar, logo "e" = "ValueError" nessa ocasiao
    print(e) #entao aqui ira printar o erro, porem com o raise utilizado na funcao media, modificamos a mensagem de erro que ira aparecer aqui
    #se acontecer esse erro a mensagem sera: "A lista não pode possuir mais de 4 notas."
else:
    print(resultado)
finally:
    print("A consulta foi encerrada!")




###--------------------------List comprehension-------------------------------------###
#forma simples e concisa de criar listas, sendo que essas listas seguirão alguns padrões, via condicionais, laços e outras expressões.
notas = [[8.0, 9.0, 10.0], [9.0, 7.0, 6.0], [3.4, 7.0, 7.0], [5.5, 6.6, 8.0], [6.0, 10.0, 9.5]]

def media(lista: list=[0]) -> float:
  ''' Função para calcular a média de notas passadas por uma lista

  lista: list, default [0]
    Lista com as notas para calcular a média
  return = calculo: float
    Média calculada
  '''
  
  calculo = sum(lista) / len(lista)

  return calculo
#aqui é apenas um for in comum
medias = [round(media(nota),1) for nota in notas] #round(media(nota),1) para arredondar para apenas uma casa decimal
#ira nos retornar a media de cada estudante (cada lista)

lista_nomes = [nome[0] for nome in nomes]
#usa uma list comprehension para criar uma nova lista (lista_nomes) contendo as iniciais dos nomes de cada aluno na lista nomes.
#a leitura do codigo ficaÇ
#Para cada elemento 'nome' (que é uma string) em lista 'nomes', extrai o primeiro caractere (nome[0]), que é a inicial do nome.
# print(lista_nomes) -->['J', 'M', 'J', 'C', 'A']

#ZIP: O zip() recebe um ou mais iteráveis (lista, string, dict, etc.) e retorna-os como um iterador de
#tuplas onde cada elemento/valores dos iteráveis são pareados/repassados.

Recebe_iteraveis_e_gera_tuplas = list(zip(nomes, medias)) #lembre-se de sempre usar a funcao list, pois se nao nos retornara <zip at 0x7fbeffc79d00
#print(Recebe_iteraveis_e_gera_tuplas) --> [('João', 9.0), ('Maria', 7.3), ('José', 5.8), ('Cláudia', 6.7), ('Ana', 8.5)]

##EXPRESSAO: valor que ira retornar dessa logica, estudante[0] = nome, entao teremos como retorno dessa logica o nome de cada looping que ter a condicao atendida
#for ITEM: geralmente utilizado como 'i' e nesse caso como 'estudante' e o nome do elemento que recebera a informacao de cada iteracao/looping) 
#in LISTA: meio autoexplicativo 'pegamos a lista estudantes'    if condicao: so serao aceitos o que atenderem a condicao, nesse caso deve estudante[1] >= 8]
#candidatos = [EXPRESSAO for ITEM in LISTA if CONDICAO] acima tem a explicacao de cada elemento e a logica
#candidatos = [estudante[0] for estudante in estudantes if estudante[1] > 7]
#print(candidatos) --> ['João', 'Ana']






nomes = [('João', 'J720'), ('Maria', 'M205'), ('José', 'J371'), ('Cláudia', 'C546'), ('Ana', 'A347')]
notas = [[8.0, 9.0, 10.0], [9.0, 7.0, 6.0], [3.4, 7.0, 7.0], [5.5, 6.6, 8.0], [6.0, 10.0, 9.5]]
medias = [9.0, 7.3, 5.8, 6.7, 8.5]
#situacao = [resultado_if if cond else resultado_else for item in lista]
situacao = ["Aprovado" if media > 5 else "Reprovado" for media in medias]
#print(situacao) --> ['Aprovado(a)', 'Aprovado(a)', 'Reprovado(a)', 'Aprovado(a)', 'Aprovado(a)']

#[expr for item in lista de listas]
cadastro = [x for x in [nomes, notas, medias]]
#print(cadastro) --> 
[ #Repare que geramos uma lista de listas com tuplas e listas
  [('João', 'J720'),
  ('Maria', 'M205'),
  ('José', 'J371'),
  ('Cláudia', 'C546'),
  ('Ana', 'A347')],
 
 [[8.0, 9.0, 10.0],
  [9.0, 7.0, 6.0],
  [3.4, 7.0, 7.0],
  [5.5, 6.6, 8.0],
  [6.0, 10.0, 9.5]],
 
 [9.0, 7.3, 5.8, 6.7, 8.5]
 ]

lista_completa = [nomes, notas, medias, situacao] 
[
[('João', 'J720'), ('Maria', 'M205'), ('José', 'J371'), ('Cláudia', 'C546'), ('Ana', 'A347')],
[[8.0, 9.0, 10.0], [9.0, 7.0, 6.0], [3.4, 7.0, 7.0], [5.5, 6.6, 8.0], [6.0, 10.0, 9.5]],
[9.0, 7.3, 5.8, 6.7, 8.5],
['Aprovado', 'Aprovado', 'Reprovado', 'Aprovado', 'Aprovado'],
]


'''
Atenção: Se testarmos rodar o código como está, teremos um problema, porque estamos considerando que
a coluna de notas vai pegar a primeira lista da lista_completa: a lista de tuplas com os nomes dos
estudantes e seus códigos.

Não é isso que queremos. Precisamos da lista de notas, depois a de média e, por fim, a de situação.
Ou seja, é necessário saltar uma linha. Para isso, vamos passar i+1.
'''

coluna = ["Notas", "Media Final", "Situação"]

#####---------Dict comprehension---------#####
cadastro  = {coluna[i]: lista_completa[i+1] for i in range(len(coluna))}

#print(cadastro) --> 
{'Notas': [[8.0, 9.0, 10.0],
  [9.0, 7.0, 6.0],
  [3.4, 7.0, 7.0],
  [5.5, 6.6, 8.0],
  [6.0, 10.0, 9.5]],
 'Média Final': [9.0, 7.3, 5.8, 6.7, 8.5],
 'Situação': ['Aprovado', 'Aprovado', 'Reprovado', 'Aprovado', 'Aprovado']}

#Por fim, vamos adicionar o nome dos estudantes, extraindo apenas seus nomes da lista de tuplas
#cadastro['Estudante'] = [expressao for item in range(len(lista_completa[0]))]
cadastro['Estudante'] = [lista_completa[0][i][0] for i in range(len(lista_completa[0]))]
#print(cadastro) --> 
{'Notas': [[8.0, 9.0, 10.0],
  [9.0, 7.0, 6.0],
  [3.4, 7.0, 7.0],
  [5.5, 6.6, 8.0],
  [6.0, 10.0, 9.5]],
 'Média Final': [9.0, 7.3, 5.8, 6.7, 8.5],
 'Situação': ['Aprovado', 'Aprovado', 'Reprovado', 'Aprovado', 'Aprovado'],
 'Estudante': ['João', 'Maria', 'José', 'Cláudia', 'Ana']}

meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
despesa = [860, 490, 1010, 780, 900, 630, 590, 770, 620, 560, 840, 360]

dicionario = {meses[i]: despesa[i] for i in range(len(meses))}

#dict comprehendion com filter
estados = ['SP', 'ES', 'MG', 'MG', 'SP', 'MG', 'ES', 'ES', 'ES', 'SP', 'SP', 'MG', 'ES', 'SP', 'RJ', 'MG',
'RJ', 'SP', 'MG', 'SP', 'ES', 'SP', 'MG']
estados_unicos = list(set(estados))
contagem_estados = {estados_unicos[i]: len(list(filter(lambda iteravel: iteravel == estados_unicos[i], estados))) for i in range(len(estados_unicos))}











#Apenas um exemplo de codigo
# Dicionário de votos por design
votos = {'Design 1': 1334, 'Design 2': 982, 'Design 3': 1751, 'Design 4': 210, 'Design 5': 1811}

# Inicializamos as variáveis
total_votos = 0 # Irá somar todos os votos 
vencedor = '' # Irá armazenar o nome do design vencedor
voto_vencedor = 0 # Irá armazenar a quantidade vencedora de votos

# Percorremos os valores de chaves e elementos do dicionário
for design, voto_desing in votos.items():
  # Somamos o total de votos
  total_votos += voto_desing
  # Verificamos se o voto do atual desing (voto_desing) é maior que o valor armazenado em voto_vencedor
  # Cada vez que voto_desing superar o valor em voto_vencedor, 
  # a variável voto_vencedor vai ser igual à voto_desing, atribuindo um novo valor
  # De forma similar, o vencedor também é substituído pelo design
  if voto_desing > voto_vencedor:
    voto_vencedor = voto_desing
    vencedor = design
# Calculamos a porcentagem do design vencedor
porcentagem = 100 * (voto_vencedor) / (total_votos)

#Resultado
print(f'{vencedor} é o vencedor: ')
print(f'Porcentagem de votos: {porcentagem}%')



































