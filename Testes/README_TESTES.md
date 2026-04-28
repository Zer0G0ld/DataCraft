# 🧪 DataCraft - Arquivos de Teste

Este diretório contém arquivos de exemplo para testar todas as funcionalidades do **DataCraft**. Todos os dados são **totalmente fictícios** e criados exclusivamente para fins de teste e validação do conversor.

---

## 📁 Arquivos Disponíveis

| Arquivo | Formato | Descrição | Registros | Finalidade |
|---------|---------|-----------|-----------|------------|
| `teste.json` | JSON | Estrutura de empresa com departamentos e funcionários | 5 funcionários | Testar JSON aninhado com listas e objetos |
| `teste.xml` | XML | Versão XML do mesmo conjunto de dados | 5 funcionários | Testar XML com tags repetitivas e hierarquia |
| `teste.yaml` | YAML | Estrutura YAML complexa com servidores | 3 servidores | Testar suporte a YAML |
| `teste.yml` | YAML | Configuração de sistema com servidores | 3 servidores | Testar arquivos .yml (alternativo) |
| `teste.csv` | CSV | Base de dados de funcionários | 20 registros | Testar arquivos tabulares com cabeçalho |
| `teste.txt` | Texto | Log de eventos no formato CSV | 10 registros | Testar arquivos .txt com estrutura CSV |

---

## 🎯 Cenários de Teste

### 1️⃣ **JSON Aninhado (`teste.json`)**
- **Estrutura**: Objeto com listas aninhadas (departamentos → funcionários → contatos → habilidades)
- **O que testa**: 
  - Detecção automática da lista principal (funcionários)
  - Achatamento de estruturas aninhadas com ponto (ex: `contact.email`)
  - Expansão de listas de valores simples (ex: `skills` como lista de strings)
  - Extração de MAC Address do campo `contact.mac`

**Exemplo de saída esperada:**
| id | name | role | contact.email | contact.mac | skills | contact.mac_MAC |
|----|------|------|---------------|-------------|--------|-----------------|
| 101 | João Santos | Senior Developer | joao.santos@techcorp.com | AA:BB:CC:DD:EE:FF | Python, JavaScript, SQL | AA:BB:CC:DD:EE:FF |

### 2️⃣ **XML Hierárquico (`teste.xml`)**
- **Estrutura**: Tags XML aninhadas com elementos repetitivos
- **O que testa**:
  - Conversão de XML para dicionário
  - Detecção de tags repetitivas (`<employee>` dentro de `<employees>`)
  - Achatamento de estruturas XML aninhadas
  - Preservação de hierarquia com pontos (ex: `contact.email`)

### 3️⃣ **YAML Configuração (`teste.yaml` e `teste.yml`)**
- **Estrutura**: Configurações com listas e objetos aninhados
- **O que testa**:
  - Suporte a formato YAML
  - Detecção automática de listas principais
  - Expansão de objetos aninhados
  - Leitura de arquivos com extensões `.yaml` e `.yml`

### 4️⃣ **CSV Tabular (`teste.csv`)**
- **Estrutura**: Planilha tradicional com cabeçalho
- **O que testa**:
  - Leitura de arquivos CSV com cabeçalho
  - Tratamento de diferentes tipos de dados (números, strings, datas)
  - Extração de MAC Address (coluna específica)
  - Suporte a acentos e caracteres especiais
  - Valores com formatação especial (telefones)

### 5️⃣ **Arquivo TXT com CSV (`teste.txt`)**
- **Estrutura**: Arquivo de log no formato CSV
- **O que testa**:
  - Detecção de formato por conteúdo (não apenas extensão)
  - Leitura de arquivos .txt estruturados
  - Extração de MAC Address de texto livre

---

## 🚀 Como Executar os Testes

### Passo 1: Inicie o DataCraft
```bash
python DataCraft.py
# ou execute o executável DataCraft.exe
```

### Passo 2: Teste cada formato
1. **Arquivo JSON**: Clique em "Buscar", selecione `teste.json`, defina saída, clique em "Converter"
2. **Arquivo XML**: Repita o processo com `teste.xml`
3. **Arquivo YAML**: Teste com `teste.yaml` e `teste.yml`
4. **Arquivo CSV**: Teste com `teste.csv`
5. **Arquivo TXT**: Teste com `teste.txt`

### Passo 3: Verifique os resultados
- O Excel gerado deve conter os dados corretamente estruturados
- Verifique se todas as colunas aninhadas foram expandidas
- Confirme que os MAC Address foram extraídos corretamente
- Valide que colunas vazias foram removidas (se opção ativada)

---

## 📊 Resultados Esperados

### Para `teste.json` e `teste.xml`:
- **Lista principal**: `employees` (5 registros)
- **Colunas esperadas**: 
  - `id`, `name`, `role`
  - `skills` (lista concatenada)
  - `contact.email`, `contact.phone`, `contact.mac`
  - `contact.mac_MAC` (MAC extraído)
  - Campos do departamento pai (podem aparecer como `department.name`, `department.manager`)

### Para `teste.yaml` e `teste.yml`:
- **Lista principal**: `servers` (3 registros)
- **Colunas esperadas**:
  - `name`, `ip`, `status`
  - `tags` (lista concatenada)
  - `config.cpu`, `config.memory`, `config.mac`
  - `config.mac_MAC` (MAC extraído)

### Para `teste.csv`:
- **Registros**: 20 linhas
- **Colunas**: `ID`, `Nome`, `Departamento`, `Cargo`, `Salario`, `Email`, `Telefone`, `MacAddress`, `DataAdmissao`, `Ativo`
- **MAC extraído**: Já existe coluna específica

### Para `teste.txt`:
- **Registros**: 10 linhas
- **Colunas**: `Timestamp`, `Level`, `Message`, `User`, `IP`, `MacAddress`
- **MAC extraído**: Coluna específica ou extraído do texto

---

## 🐛 Relatando Problemas

Se algum teste falhar ou apresentar comportamento inesperado:

1. **Capture uma screenshot** do erro ou resultado incorreto
2. **Copie o log** da área de log do DataCraft
3. **Abra uma issue** no GitHub com:
   - Descrição do problema
   - Qual arquivo de teste falhou
   - Log completo
   - Sistema operacional e versão do DataCraft

---

## 📝 Notas Importantes

### ⚠️ Dados Fictícios
Todos os dados contidos nestes arquivos são **totalmente fictícios**:
- Nomes de pessoas, empresas e locais são inventados
- Endereços IP, MAC Address e contatos são aleatórios
- Qualquer semelhança com a realidade é mera coincidência

### 🔧 Recomendações
- **Nunca use dados reais** para testes iniciais
- **Sempre verifique** os resultados antes de usar em produção
- **Faça backup** de seus arquivos originais antes de processar

### 📌 Funcionalidades Testadas
- [x] Detecção automática de formato
- [x] Processamento de JSON aninhado
- [x] Processamento de XML hierárquico
- [x] Suporte a YAML
- [x] Leitura de CSV com cabeçalho
- [x] Extração automática de MAC Address
- [x] Achatamento de estruturas aninhadas
- [x] Remoção de colunas vazias
- [x] Geração de aba resumida
- [x] Auto-abrir arquivo após conversão

---

## 🎯 Checklist de Testes

Use esta checklist para validar todas as funcionalidades:

### JSON (`teste.json`)
- [ ] Abre o arquivo sem erros
- [ ] Detecta lista principal automaticamente
- [ ] Expandir estruturas aninhadas funciona
- [ ] MAC Address extraído corretamente
- [ ] Gera planilha com todos os dados
- [ ] Aba resumida é criada

### XML (`teste.xml`)
- [ ] Converte XML para dicionário corretamente
- [ ] Detecta tags repetitivas
- [ ] Achatamento mantém hierarquia
- [ ] Dados são salvos corretamente

### YAML (`teste.yaml` / `teste.yml`)
- [ ] Arquivo .yaml carrega corretamente
- [ ] Arquivo .yml carrega corretamente
- [ ] Estrutura YAML é processada

### CSV (`teste.csv`)
- [ ] Cabeçalho é lido corretamente
- [ ] Todos os 20 registros são processados
- [ ] Tipos de dados são preservados
- [ ] Acentos são mantidos

### TXT (`teste.txt`)
- [ ] Arquivo .txt é reconhecido como CSV
- [ ] Formato é detectado pelo conteúdo
- [ ] Dados são extraídos corretamente

---

## 📚 Mais Informações

- **Documentação principal**: [README.md](../README.md)
- **Reportar bugs**: [Issues](https://github.com/Zer0G0ld/DataCraft/issues)
- **Autor**: Zer0G0ld

---

**DataCraft** - Transforme dados em ouro 🏆

*Última atualização: Março 2026*

## 📁 **Estrutura de pastas sugerida:**

```
DataCraft/
├── DataCraft.py
├── README.md
├── README_TESTES.md        # <-- Este arquivo
├── LICENSE
├── requirements.txt
├── voto.ico
├── img/
│   └── DataCraft.PNG
└── testes/                  # <-- Pasta com os arquivos de teste
    ├── teste.json
    ├── teste.xml
    ├── teste.yaml
    ├── teste.yml
    ├── teste.csv
    └── teste.txt
```