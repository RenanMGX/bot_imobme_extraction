# Descrição de Uso do Sistema de Extração de Relatórios Imobme

## Sobre
O sistema automatiza a extração e tratamento de relatórios da plataforma Imobme, focado em gestão imobiliária. Utiliza Selenium para automação web, permitindo a extração de dados diretamente da interface da plataforma. Além disso, manipula arquivos Excel para tratamento dos dados extraídos e emprega SFTP para transferência segura de arquivos.

## Componentes
- `SambaExtract.py`: Responsável pela extração de dados via SMB.
- `Financeiro.py`: Trata especificamente dos dados financeiros extraídos.
- `IntegracaoWebExtract.py`, `extraction_imobme.py`: Automatizam a navegação e extração de dados da plataforma Imobme.
- `tratar_arquivos_excel_imobme.py`: Manipula e trata os dados extraídos, organizando-os em arquivos Excel.
- `sftp.py`: Gerencia a transferência segura de arquivos via SFTP.
- `registro.py`, `carregar_credenciais.py`: Auxiliam no registro de logs e no carregamento de credenciais necessárias para a execução do sistema.
- `main.py`: Arquivo principal que coordena a execução dos componentes do sistema.
- `index.html`: Interface de usuário para interação com o sistema (se aplicável).

## Funcionamento
1. **Extração de Dados**: O sistema navega pela plataforma Imobme, extrai relatórios especificados e salva os dados localmente.
2. **Tratamento de Dados**: Após a extração, os dados são tratados e organizados em arquivos Excel para análise e relatórios.
3. **Transferência de Arquivos**: Os arquivos tratados são então transferidos via SFTP para um servidor ou local seguro especificado.

## Requisitos
- Python 3.x
- Bibliotecas Python: `selenium`, `pandas`, `openpyxl`, `paramiko`, entre outras.
- WebDriver para o navegador escolhido (e.g., ChromeDriver para Google Chrome).

## Configuração e Execução
1. Instale as dependências Python necessárias.
2. Configure as credenciais e parâmetros necessários em `carregar_credenciais.py`.
3. Execute `main.py` para iniciar o processo de extração e tratamento dos dados.

## Notas
- Certifique-se de ter as permissões necessárias na plataforma Imobme para acessar e extrair os dados desejados.
- A configuração do WebDriver deve corresponder à versão do navegador utilizado.
