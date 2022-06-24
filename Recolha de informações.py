import platform, socket, os, wmi, psutil as p, shutil, win32api, multiprocessing, time
from cpuinfo import get_cpu_info
from psutil import virtual_memory
from uuid import getnode as get_mac
import pyodbc # Linhas 1 a 5, vamos importar as bibliotecas necessárias para o programa funcionar corretamente

def job():

  conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=srvsql-ipt.ddns.net;DATABASE=DB_81745_PA;UID=81745;PWD=81745')  
  cursor = conn.cursor() #Linhas 9 e 10 vão servir para fazer uma ligação correta à base de dados


  def size(byte):
    for x in ["B","KB","MB","GB","TB"]:
      if byte<1024:
        return f"{byte:.2f}{x}"
      byte=byte/1024 #Das linhas 13 até à 17, Vai servir para o programa mostrar a ram em GB (neste caso), em vez de mostrar o total de ram em bytes



  mac = get_mac() # Vai buscar o enderenço mac

  info = get_cpu_info() # Vai buscar informações do cpu

  uname = platform.uname() # Vai buscar várias informações do pc

  processador = info.get('brand_raw') # Vai buscar o modelo exato do cpu

  ip = socket.gethostbyname(socket.gethostname()) # Vai buscar o enderenço ip do pc

  so = platform.system() # Vai buscar o nome do Sistema Operativo Ex: Windows

  vso = platform.release() # Vai buscar o número da versão do sistema operativo Ex: 10 (de windows 10)

  vaso = platform.version() # Vai buscar a versão do sistema operativo

  nc = uname.node # Vai buscar o nome do computador

  nu = os.getlogin() # Vai buscar o nome do utilizador Ex: vasco

  ac = uname.machine # Vai buscar a arquitetura do computador Ex: 64 bits

  f = wmi.WMI() # define que a letra f Vai ser a variável wmi.WMI()

  ccpu = multiprocessing.cpu_count() # Vai buscar o número de cores do processador

  utcpu =  p.cpu_percent(4) # Vai buscar a utilização do processador por 4 segundos

  gpu_info = f.Win32_VideoController()[0] # Vai buscar várias informações da placa gráfica

  outputgpu = format(gpu_info.Name) # Vai buscar o nome da placa gráfica

  mem = p.virtual_memory() # Vai buscar várias informações sobre a ram

  ram = mem.total #Vai buscar o total de ram instalada no PC

  total, usado, livre = shutil.disk_usage("/") # Vai buscar várias informações relativas aos espaço do disco

  et = (total // (2**30)) #Vai meter o espaço total em GB

  eu = (usado // (2**30)) #Vai meter o espaço usado em GB

  el = (livre // (2**30)) #Vai meter o espaço livre em GB

  path = "C:/"

  info = win32api.GetVolumeInformation(path)

  nsd = info[1] # linhas 64 à 68 servem para configurar em que disco Vai buscar as informações e depois guardar numa variável o número de série (neste caso sendo do disco C)

  gpus = f.Win32_VideoController()[0] #Vai buscar as placas gráficas (GPUS) disponíveis no pc

  et_bd = str(et) #converte para string

  eu_bd = str(eu) #converte para string

  el_bd = str(el) #converte para string

  nsd_bd = str(nsd) #converte para string

  ip_bd = str(ip) #converte para string

  mac_bd = str(mac) #converte para string

  gpu_bd = str(outputgpu) #converte para string

  utcpu_bd = str(utcpu) #converte para string

  ccpu_bd = str(ccpu) #converte para string

  cursor.execute("SELECT * FROM specs_net WHERE ma = '%s'" % mac)

  results = cursor.fetchall()

  if results:
      print("Você já guardou as suas informações na base de dados, por isso o programa Vai atualizar as suas informações")
      sqlupdatequery = "UPDATE specs_pc SET so =  '"+so+"', vso =  '"+vso+"' , vaso =  '"+vaso+"' , nomepc = '"+nc+"', nomeuti = '"+nu+"', ac = '"+ac+"', cpu = '"+processador+"', gpu = '"+gpu_bd+"', ram = '"+size(ram)+"', et = '"+et_bd+"', eu = '"+eu_bd+"', el = '"+el_bd+"',nsd = '"+nsd_bd+"', ccpu =  '"+ccpu_bd+"', utcpu =  '"+utcpu_bd+"' WHERE ma_id ='%s'" % mac
      cursor.execute(sqlupdatequery)
      cursor.commit()
      cursor.close() # Da lina 91 até à 100, se o mac address que a variável "mac" receber for igual a algum dos mac addresses da base de dados, o programa irá somente atualizar as informações desse computador. Caso contrário irá registar um novo computador na base de dados
      
  else:
    print("Bem-vindo ao nosso programa") 

    print(f"Sistema Operativo: {so} ") # Vai dar output do nome e versão do so Ex: Windows 11

    print(f"Nome do computador: {nc}") # Vai dar output do nome do computador

    print(f"Nome do utilizador: {nu}") # Vai dar output do nome do utilizador Ex: Vasco

    print(f"Arquitetura do computador: {ac}") # Vai dar output do tipo de arquitetura do computador Ex: 64 bits

    print(f"Processador: {processador}") # Vai dar output do modelo exato do processador

    print(f"Núcleos do processador: {ccpu}") # Vai dar output do número de cores do processador

    print(f"Utilização do cpu: {utcpu}") # Vai dar output da utilização do processador Ex: 55.3

    print(f'Gpu: {outputgpu}') # Vai dar output das placas gráficas

    print(f"Endereço IP: {ip}") # Vai dar output do endereço ip

    print(f"Endereço MAC: {mac}") # Vai dar output do endereço MAC

    print("Ram: ",size(ram)) # Vai dar output do total de ram instalada no pc

    print("Espaço total do Disco Rigido:", et, "GB") # Vai dar output do espaço total do disco rigido

    print("Espaço usado do Disco Rigido:", eu, "GB") # Vai dar output do espaço usado do disco rigido

    print("Espaço livre do Disco Rigido:", el, "GB" ) # Vai dar output do espaço livre do disco rigido

    print( "Número de série do disco:", nsd) # Vai dar output do número de série do disco C

    sqlquery = "INSERT INTO specs_pc(so, vso, vaso, nomepc, nomeuti, ac, cpu, gpu, ram, et, eu, el,nsd, ma_id, ccpu, utcpu) VALUES ('"+so+"', '"+vso+"', '"+vaso+"', '"+nc+"', '"+nu+"', '"+ac+"', '"+processador+"', '"+gpu_bd+"', '"+size(ram)+"', '"+et_bd+"', '"+eu_bd+"', '"+el_bd+"', '"+nsd_bd+"','"+mac_bd+"', '"+ccpu_bd+"', '"+utcpu_bd+"')"
    sqlquery2 = "INSERT INTO specs_net(ip,ma) VALUES ('"+ip_bd+"', '"+mac_bd+"')"
    cursor.execute(sqlquery2)
    cursor.execute(sqlquery)

    cursor.commit()

    print("Dados guardados com sucesso")

    cursor.close() # Das linhas 135 até à 144, vamos fazer os inserts de todas as informações do computador (caso as informações sejam novas). Caso os dados sejam inseridos corretamente, o programa irá mostrar a mensagem "Dados guardados com sucesso".

if __name__ == '__main__':
  while True:
    job()
    time.sleep(5)    