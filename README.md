# Enviando vários e-mails com python 📧


<h2>Iniciando o projeto</h2>

O primeiro passo é instalar as bibliotecas para utilizar no projeto, abra o seu terminal e digite os seguintes comandos:

Pywin32com:

![lib1](https://user-images.githubusercontent.com/77951123/190929592-4f6b40b8-2d4b-4849-ba91-63f68c09f312.gif)


Pandas:

![lib2](https://user-images.githubusercontent.com/77951123/190929595-32f9028c-7045-4c3c-a4c7-b5010f2f063d.gif)



Depois vamos criar uma pasta e dentro dela um arquivo .py com o nome da sua preferência; (use o editor de código que você mais gosta =D)

![image](https://user-images.githubusercontent.com/77951123/190929669-571a0d6c-6a9b-402f-b806-1e09885bea37.png)

No meu caso criei um arquivo .py e outro .ipynb que é um notebook do jupyter para separar melhor o código; 

  <hr>
  
<h2>Importanto a Biblioteca</h2>

Dentro do arquivo vamos começar importando a biblioteca pywin32 que te permite automatizar uma série de coisas no Windows, então isso pode te facilitar bastante caso utilize o mesmo e o 
Pandas que fornece ferramentas para análise e manipulação de dados.

![libs](https://user-images.githubusercontent.com/77951123/190929751-c258698f-ce1b-4b70-ab5c-ecb40798f3fd.gif)

<hr>

<h2>Lendo um arquivo Excel com Pandas</h2>

Utilizaremos o Pandas para ler um arquivo excel com todos os emails que precisamos, se o arquivo que deseja ler está dentro da pasta do código você só precisa colocar o seu nome e extensão entre aspas ("seuArquivo.xlsx") mas se não for esse o seu caso apenas cole o <b>path</b> do seu arquivo;


![lendoexcel](https://user-images.githubusercontent.com/77951123/190929886-628862e6-8e51-4796-bb3e-897706a88235.gif)

<hr>


## Convertendo a Coluna para lista
Para o nosso código funcionar vamos precisar converter a coluna que desejamos para o tipo list;

Váriavel = list(tabela['NomeColuna'])



![listando](https://user-images.githubusercontent.com/77951123/190930140-56cef370-bbf0-4908-9479-c63c61556634.gif)

<hr>

## Convertendo a lista para string

E por último precisamos converter a lista para o tipo string e tratar os seus dados com o método replace;


![convert](https://user-images.githubusercontent.com/77951123/190930311-2682d3d4-99b8-4252-8a85-2885da6e16de.gif)

<hr>

# Iniciando o Email

## Definindo os objetos

Apenas precisamos definir que objeto vamos utilizar que no caso é o OUTLOOK e logo depois precisamos definir a váriavell mais importante que é a <b>mailto</b> utilizando a váriavel que criamos acima do tipo string;

![image](https://user-images.githubusercontent.com/77951123/190930397-5bb8f03b-aa09-45ef-a21a-00f95c6cfd54.png)

## Lidando com várias contas dentro do outlook

Caso tenha várias contas logadas no seu outlook mas você quer apenas utilizar uma para enviar seus email precisamos criar uma estrutura de FOR:

![image](https://user-images.githubusercontent.com/77951123/183557633-20f9f0f5-1c53-4b07-bdbb-dbcfc883e579.png)

E agora é só criar o objeto e definir um IF;

![image](https://user-images.githubusercontent.com/77951123/190930587-3bd868cc-6561-439f-a64d-9c677ab069e3.png)

<hr>

# Explicando cada campo para edição

 Depois de ter definido o objeto e o IF agora só devemos preencher os seguintes itens;
 
<b>mail.to</b>: Aqui vamos colocar a váriavel que criamos <b>mailto</b> para que o email seja enviado para todos os contatos que definimos;

![image](https://user-images.githubusercontent.com/77951123/190930643-6f068716-6d7f-407e-bfca-ca0afa8e7542.png)


<b>mail.Subject</b>: É onde devemos colocar o assunto;

![image](https://user-images.githubusercontent.com/77951123/190930694-7ead6f18-e149-435e-9ec8-a876f5615f50.png)

<b>mail.CC e mail.BCC</b>: Envia uma cópia do email e BCC envia uma cópia oculta do mesmo;

![image](https://user-images.githubusercontent.com/77951123/183558669-5a4b2311-4ea4-4631-9b58-ec8338dcf6f1.png)

<b>mailHTMLBody</b>: É onde a mágica acontece aqui definimos a mensagem que vamos enviar e a melhor forma de fazer isso é utilizando o HTML como no exemplo abaixo:

![image](https://user-images.githubusercontent.com/77951123/183558888-9095567a-4587-44ef-887a-b78b69eed5aa.png)

Mas relaxa vc pode usar o mail.Body caso ainda não saiba muito sobre HTML!

# Enviando Anexo

Para enviar anexo precisamos apenas definir uma váriavel com o caminho do arquivo e enviar;

![image](https://user-images.githubusercontent.com/77951123/190931140-837579c9-f0e5-4fbe-9604-6159ff034a38.png)



















