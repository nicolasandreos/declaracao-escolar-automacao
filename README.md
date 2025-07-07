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

---

## 📦 Requisitos

- Python 3.8+
- Word instalado (para impressão funcionar corretamente)

### 📚 Bibliotecas necessárias:

Instale com:

```bash
pip install python-docx pywin32
```
🛠️ Como usar
Prepare o arquivo nomes.txt com um nome por linha:

Copy
Edit
João da Silva
Maria Oliveira
Carlos Souza
Prepare o modelo Declaração Escolar.docx com os seguintes placeholders:

less
Copy
Edit
Declaramos para os devidos fins que [NOME] está regularmente matriculado na instituição.
Cidade: [CIDADE], [DIA] de [MES] de [ANO].
Execute o script:

bash
Copy
Edit
python script.py
Digite a cidade quando solicitado.

Os arquivos .docx serão gerados automaticamente na pasta Declaracoes/.

🖨️ Impressão automática (opcional)
Após gerar os arquivos, o script oferece a opção de imprimir automaticamente os documentos utilizando a impressora selecionada pelo usuário.

⚠️ Necessário estar no Windows com o Microsoft Word instalado.

