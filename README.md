# HR Automation – *WhatsApp* WebScrapping
---
> #### **Objetivo**: automação para identificar colaboradores que iniciarão as atividades no dia atual e enviar mensagens, personalizadas de acordo com nome e genêro de cada funcionário, e arquivo de vídeo. Através da plataforma online do aplicativo *WhatsApp*.
<br>

> #### **Arquivos disponibilizados neste repositório**: 
> - Códigos de 3 arquivos VBA:
>> - Código aplicado ao livro de Excel para permitir ativação do macro quando a planilha é ativada e verificar se o envio das mensagens já ocorreu naquele dia.
>> - Módulo para atualizar tabela diária de colaboradores contratados que estão iniciando, a partir da tabela geral de contratações da empresa. 
>> - Módulo para montar mensagens de boas-vindas, montar link URL para [API Whatsapp](api.whatsapp.com) com a mensagem personalizada e número de telefone do colaborador, abrir o Google Chrome (via Selenium), conectar ao Web Whastapp via cookies, enviar mensagem e arquivo de vídeo para cada colaborador da tabela diária, registrar que as mensagens foram enviadas naquele dia e caso algum envio falhe.
> #### **A planilha utilizada na execução do Macro** não será adicionada ao repositório em respeito as normas LGPD em relação aos dados da empresa em questão.
<br>

> #### **Pré-Requisitos**:
> - Pasta de Trabalho do Microsoft Excel Habilitado Para Macros (arquivo *.xlsm*)
> - Selenium Basic v2.0.9.0 com ChromeDriver atualizado

## **Por**: Guilherme Moret
## **Empresa**: Grupo Rovema – Porto Velho, Rondônia, Brasil

