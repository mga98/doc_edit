<h1 align="center"> Draw-Project </h1>

<p align="center">
<img src="http://img.shields.io/static/v1?label=STATUS&message=EM%20DESENVOLVIMENTO&color=GREEN&style=for-the-badge"/>
</p>

<h2> Descrição </h2>

<p>
    doc_edit é um script que permite o usuário preencher documentos .docx (Word)
    automaticamente através de campos de um formulário .json ou por argumentos em uma função. O script fornece uma criação dos campos dos formulários terminados com um underscore (ex: nome_). Com isso o usuário precisa apenas abrir o documento a ser editado e preencher os campos que querem substituir com a mesma sintaxe. Exemplo em
    <a href="https://github.com/mga98/doc_edit/blob/main/interface.py">interface.py</a>

<div align='center'>

![Exemplo campos](https://user-images.githubusercontent.com/95861523/219868755-fed5fa5d-75a3-4e1e-92fc-2ec8f67237e7.png)

</div>

</p>

<h2> Funções </h2>

<h3>open_document()</h3>
<p>Abre um documento .docx (Word) que você deseja preencher.</p>

<h3>create_form()</h3>
<p>Cria um formulário .json com as chaves que você utilizará depois para preencher o documento.</p>

<h3>upload_form()</h3>
<p>Abre um formulário .json já criado com as chaves a serem preenchidas.</p>

<h3>update_document()</h3>
<p>Gera um documento .docx com as chaves do formulário preenchidas. Ou as chaves e os valores podem ser passados como parâmetros em formato de dicionário.</p>

<h2> Tecnologias utilizadas </h2>

<ul>
<li>Python</li>
<li>python-docx</li>
<li>json</li>
</ul>

<h2> Rodando o projeto </h2>
<h4> Dependências </h4>
<ul>
<li>Python 3.0 ou +</li>
<li>python-docx</li>
</ul>
<h4> Clonando o projeto </h4>

```
git clone git@github.com:mga98/doc_edit.git
```

<h4> Executando o projeto </h4>
<p> Abra o terminal com o ambiente virtual ativo e execute os seguintes comandos: </p>

```
pip install -r requirements.txt
python interface.py
```
