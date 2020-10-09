import pandas as pd 

#comandos
#pipenv install --dev
#pipenv run python3 resolucaoIago.py
#Autor: Iago Costa das Flores
#Processo Seletivo SICOM
class Servidor:
    
    def __init__(self, Nome_Servidor, Matricula_Servidor, Arquivo_Servidor):
        self.Nome_Servidor = Nome_Servidor
        self.Matricula_Servidor = Matricula_Servidor
        self.Arquivo_Servidor = Arquivo_Servidor
        self.Dataframe = pd.read_excel(self.Arquivo_Servidor, index_col=None, dtype={'Nº do Registro':str, 'Nome':str, 'Data':str, 'Atividades':str, 'Endereço':str, 'Bairro':str, 'Servidor':str, 'Procedimentos':str}) #lendo dados do excel
        self.Dataframe = self.Dataframe.dropna() #para limpar valores vazios
    
#implementação do CRUD 
    def CarregarShape(self):
        return self.Dataframe #carrega os Dados

    def CarregarLinha(self,id):
        return self.Dataframe.loc[id] 

    def DeletarLinha(self, id):
        buffer = self.Dataframe.drop(id) 
        self.Dataframe = buffer.reset_index() #reseta os indices

    def UpdateLinha(self, linha, coluna, valor):
        buffer = self.Dataframe.replace(to_replace=self.Dataframe.loc[linha,coluna], value=valor)
        self.Dataframe = buffer
    
    def InserirLinha(self, linha):
        buffer = self.Dataframe.append(linha, ignore_index=True)
        self.Dataframe = buffer.reset_index()


obj = Servidor('ServidorSICOM', '001', 'RelatorioPS.xlsx') #carregando o Objeto Servidor

#print(obj.CarregarShape())
#print(obj.CarregarLinha(509))

opcao = 'Null'
while opcao != 'S':
    opcao = str(input("O que deseja Fazer? (Obs.:Digite sem as aspas) \n 1-Digite 'C' para Criar nova Linha. \n 2-Digite 'R' para Ler uma Linha. \n 3-Digite 'U' para atualizar uma linha. \n 4-Digite 'D' para Deletar uma linha. \n 5-Digite 'Save' para Salvar Alterações \n 6-Digite 'Ver' para ver todas as linhas  \n 7-Digite 'S' para sair. \n")).upper()
    if opcao == 'C':
        Numero_Registro = str(input('Qual o número do Registro?')).upper()
        Nome = str(input("Qual o Nome?")).upper()
        Data = str(input("Qual a data?")).upper()
        Atividade = str(input("Qual a Atividade?")).upper()
        Endereco = str(input("Qual o endereço?")).upper()
        Bairro = str(input("Qual Bairro?")).upper()
        Servidor = str(input("Qual o Servidor?")).upper()
        Procedimentos = str(input("Quais os Procedimentos?")).upper()
        linha = {'Nº do Registro':Numero_Registro, 'Nome':Nome, 'Data':Data, 'Atividades':Atividade, 'Endereço':Endereco, 'Bairro':Bairro, 'Servidor':Servidor,'Procedimentos':Procedimentos}
        obj.InserirLinha(linha)
    
    if opcao == 'R':
        Numero_da_Linha_Ler = int(input('Qual a linha deseja ver?'))
        print(obj.CarregarLinha(Numero_da_Linha_Ler))

    if opcao == 'U':
        Numero_da_Linha_Update = int(input('Qual a linha deseja alterar?'))
        Numero_Coluna_da_Linha = int(input("O que deseja Fazer? (Obs.:Digite sem as aspas) \n 1-Digite '1' para Editar Numero de Registro. \n 2-Digite '2' para Editar Nome. \n 3-Digite '3' para Editar Data. \n 4-Digite '4' para Editar Atividade. \n 5-Digite '5' para Editar Endereço \n 6-Digite '6' para Editar Bairro \n 7-Digite '7' para editar Servidor \n 8-Digite '8' para editar Procedimentos \n"))    
        if Numero_Coluna_da_Linha == 1: Coluna_da_Linha = 'Nº do Registro'
        if Numero_Coluna_da_Linha == 2: Coluna_da_Linha = 'Nome'
        if Numero_Coluna_da_Linha == 3: Coluna_da_Linha = 'Data'
        if Numero_Coluna_da_Linha == 4: Coluna_da_Linha = 'Atividades'
        if Numero_Coluna_da_Linha == 5: Coluna_da_Linha = 'Endereço'
        if Numero_Coluna_da_Linha == 6: Coluna_da_Linha = 'Bairro'
        if Numero_Coluna_da_Linha == 7: Coluna_da_Linha = 'Servidor'
        if Numero_Coluna_da_Linha == 8: Coluna_da_Linha = 'Procedimentos'
        Novo_Valor = str(input('Qual o novo valor?')).upper()
        obj.UpdateLinha(Numero_da_Linha_Update, Coluna_da_Linha, Novo_Valor)
    
    if opcao == 'D':
        Numero_da_Linha_Deletar = int(input('Qual a linha deseja deletar? \n'))
        obj.DeletarLinha(Numero_da_Linha_Deletar)
        print("Linha deletada com sucesso! \n")
    
    if opcao == 'SAVE':
        excel = pd.ExcelWriter('ServidorAtualizado.xlsx', engine='xlsxwriter') #cria novo excel
        obj.Dataframe.to_excel(excel, sheet_name='Dados Atualizados') #escreve os dados do dataframe no excel
        excel.save() #salva o arquivo

    if opcao == 'VER':
        print(obj.CarregarShape())