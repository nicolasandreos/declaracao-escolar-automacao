# 📄 Automação de Declarações Escolares

Este projeto automatiza a geração de declarações escolares personalizadas em arquivos `.docx`, utilizando um modelo base e uma lista de nomes. As declarações geradas são salvas automaticamente e podem ser enviadas para impressão em lote.

---

## ✅ Funcionalidades

- Substituição automática de placeholders no documento Word:
  - `[NOME]`, `[CIDADE]`, `[DIA]`, `[MES]`, `[ANO]`
- Aplicação de estilo no nome (negrito e sublinhado).
- Geração de um documento `.docx` por pessoa.
- Impressão automática de todos os documentos gerados.
- Suporte a nomes de mês em **português** com base na localidade do sistema.

---

## 📦 Requisitos

- Python 3.8+
- Word instalado (para impressão funcionar corretamente)

## 📚 Bibliotecas necessárias:

Instale com:

```bash
pip install python-docx pywin32
```

Ótimo ponto! Essa seção realmente pode ser mais clara, fluida e bem formatada.

Aqui está a versão **melhorada e mais profissional** da seção **“Como usar”**, com uma linguagem direta, visualmente agradável e sem repetições confusas:

---

## 🛠️ Como usar

1. **Crie uma nova pasta** e abra em uma IDE (recomendado: [Visual Studio Code](https://code.visualstudio.com/)).

2. **Crie um ambiente virtual (opcional, mas recomendado):**

```bash
python -m venv venv
```

Ative o ambiente virtual:

* No **Windows** (cmd ou PowerShell):

```bash
venv\Scripts\activate
```

> ⚠️ Se a execução estiver desabilitada no PowerShell, execute:
>
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

3. **Clone este repositório (requer o Git instalado):**

```bash
git clone https://github.com/nicolasandreos/declaracao-escolar-automacao
cd declaracao-escolar-automacao/AutomacaoDeclaracoes
```

4. **Instale as dependências do projeto:**

```bash
pip install -r requirements.txt
```

5. **Abra o arquivo `nomes.txt`**

Liste um nome por linha, por exemplo:

```
João da Silva
Maria Oliveira
Carlos Souza
```

6. **Prepare o modelo Word**

O projeto já usa um arquivo chamado `Declaração Escolar.docx` com todos os **placeholders** necessários, como:

* `[NOME]`
* `[CIDADE]`
* `[DIA]`
* `[MES]`
* `[ANO]`

Esses placeholders serão automaticamente **substituídos pelos valores reais** durante a geração dos documentos.

> 💡 **Quer mudar o layout do documento?**
> Basta abrir o `Declaração Escolar.docx` no Word e mover os placeholders para onde quiser, ou adicionar/remover conforme necessário — apenas mantenha os nomes entre colchetes, como `[NOME]`.

7. **Execute o script**

No terminal, rode:

```bash
python script.py
```

Você será solicitado a digitar a cidade da instituição. O restante (data, mês, ano) será preenchido automaticamente com base na data atual do sistema.

8. **Resultado**

Os arquivos `.docx` personalizados serão salvos automaticamente na pasta `Declaracoes/`, prontos para envio ou impressão.

