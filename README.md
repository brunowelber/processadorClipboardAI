# Processador de Clipboard com IA para NVDA

Este é um addon para o leitor de telas NVDA que utiliza o poder da inteligência artificial (através da API da OpenAI) para processar e transformar textos diretamente do seu clipboard ou do texto que você selecionou.

**Autor:** Bruno Welber

## Funcionalidades

- **Processamento Inteligente:** Envie o texto da sua área de transferência ou a seleção atual para a API da OpenAI para realizar diversas tarefas.
- **Múltiplas Ações (Prompts):** Vem com um conjunto de prompts padrão, como:
  - Melhorar Ortografia e Gramática
  - Traduzir para Inglês
  - Resumir em Pontos-Chave
  - Tornar o texto mais formal (corporativo)
- **Resultado no Clipboard:** Após o processamento, o texto modificado pela IA é automaticamente copiado para a sua área de transferência, pronto para ser colado.
- **Totalmente Configurável:** Adicione, edite e remova seus próprios prompts personalizados para adaptar o addon às suas necessidades.

## Teclas de Atalho

- **Processar Texto Selecionado:** `NVDA+Shift+P`
  - Pega o texto que está atualmente selecionado, abre um menu para você escolher o prompt desejado e envia para processamento.

- **Processar Texto da Área de Transferência:** `Control+NVDA+Shift+P`
  - Pega o texto que está na sua área de transferência (clipboard), abre um menu para você escolher o prompt e envia para processamento.

Após a conclusão, um som de "beep" será emitido e o texto processado estará disponível na sua área de transferência.

## Instalação

1. Baixe o arquivo `processadorClipboardAI.nvda-addon`.
2. Abra o arquivo baixado. O NVDA perguntará se você deseja instalar o addon.
3. Confirme a instalação e reinicie o NVDA quando solicitado.

## Configuração

Para que o addon funcione, você precisa configurar sua chave de API da OpenAI.

1. Vá para o menu do NVDA (`NVDA+N`).
2. Navegue até `Preferências` > `Configurações`.
3. Na lista de categorias, encontre e selecione `Processador de Clipboard com IA`.
4. No painel de configurações, você encontrará as seguintes opções:

### Configurações Gerais

- **Chave da API:** Insira sua chave de API da OpenAI aqui. Você pode obter uma no [site da OpenAI](https://platform.openai.com/api-keys). O addon também pode usar a chave se ela estiver definida na variável de ambiente `OPENAI_API_KEY`.
- **Modelo:** Escolha o modelo de IA que deseja usar (ex: `gpt-4o`, `gpt-4-turbo`).
- **Prompt Padrão:** Selecione o prompt que aparecerá pré-selecionado no menu de escolha.

### Gerenciador de Prompts

Esta seção permite que você personalize as ações de IA:

- **Lista de Prompts:** Mostra todos os prompts disponíveis. Selecione um para ver ou editar seu conteúdo.
- **Conteúdo do Prompt:** Caixa de texto onde você pode visualizar e editar a instrução que será enviada para a IA.
- **Botões:**
  - **Salvar Edição:** Salva as alterações feitas no conteúdo de um prompt selecionado.
  - **Excluir:** Remove o prompt selecionado.
  - **Adicionar...:** Abre uma nova janela para você criar um novo prompt com um nome e uma instrução personalizados.

**Importante:** Após fazer qualquer alteração, clique em `Aplicar` ou `OK` para que as configurações (incluindo os prompts personalizados) sejam salvas permanentemente. Os prompts são salvos no arquivo `prompts.ini` dentro da pasta do addon.
