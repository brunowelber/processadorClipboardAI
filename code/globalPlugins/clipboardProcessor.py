# -*- coding: utf-8 -*-

import globalPluginHandler
import scriptHandler
import ui
import wx
import threading
import json
from pathlib import Path
import sys
import api
import configparser
import os
import tones
import copy
import re
import mimetypes
import base64
import tempfile
from urllib.parse import urlparse
from html.parser import HTMLParser

# Adiciona a pasta 'lib' ao path para encontrar dependências empacotadas
try:
    addon_dir = Path(__file__).parent
    lib_path = str(addon_dir / "lib")
    if lib_path not in sys.path:
        sys.path.append(lib_path)
    import requests
except ImportError:
    requests = None

import config
from gui import guiHelper, settingsDialogs

# --- Lógica de Prompts ---
PROMPTS_INI_PATH = Path(__file__).parent / "prompts.ini"
try:
    USER_PROMPTS_PATH = Path(config.getUserDefaultConfigPath()) / "clipboardProcessor_prompts.ini"
except Exception:
    # Fallback defensivo caso a API do NVDA não esteja disponível.
    USER_PROMPTS_PATH = Path.home() / ".clipboardProcessor_prompts.ini"
DEFAULT_PROMPT_NAME = "Melhorar Ortografia e Gramática"
INVALID_PROMPT_NAME_CHARS = {"]", "\n", "\r"}
DEFAULT_PROMPTS = {
    DEFAULT_PROMPT_NAME: {"prompt": "Sua tarefa é refinar o texto a seguir. Mantenha o estilo e a voz originais do autor, mas melhore a clareza, a coesão e a concisão. Corrija todos os erros de ortografia e gramática. Elimine a prolixidade e torne a escrita mais direta e polida.", "model": None},
    "Traduzir para Inglês": {"prompt": "traduza o seguinte texto para o inglês.", "model": None},
    "Resumir em Pontos-Chave": {"prompt": "Resuma o texto a seguir em uma lista de pontos-chave.", "model": None},
    "Tornar Mais Formal": {"prompt": "Reescreva o texto a seguir em um tom mais formal e profissional.", "model": None},
}

def _normalize_prompt_entry(value):
    if isinstance(value, dict):
        prompt_text = value.get("prompt", "")
        model = value.get("model") or None
    else:
        prompt_text = str(value)
        model = None
    return {"prompt": prompt_text, "model": model}

def _read_prompts_from(path):
    prompts = {}
    parser = configparser.RawConfigParser()
    parser.read(path, encoding='utf-8')
    for section in parser.sections():
        if 'prompt' in parser[section]:
            prompts[section] = _normalize_prompt_entry({
                "prompt": parser[section]['prompt'],
                "model": parser[section].get("model", "").strip() or None
            })
    return prompts

def load_prompts():
    # Preferir prompts do usuário (persistem entre atualizações); caso não existam, usar os do pacote.
    if USER_PROMPTS_PATH.exists():
        return _read_prompts_from(USER_PROMPTS_PATH)

    if PROMPTS_INI_PATH.exists():
        prompts = _read_prompts_from(PROMPTS_INI_PATH)
        save_prompts(prompts)  # sincroniza para o armazenamento do usuário
        return prompts

    save_prompts(DEFAULT_PROMPTS)
    return DEFAULT_PROMPTS

def save_prompts(prompts_dict):
    # RawConfigParser evita erro de interpolação ao salvar textos com %.
    parser = configparser.RawConfigParser()
    for name, value in prompts_dict.items():
        entry = _normalize_prompt_entry(value)
        parser[name] = {'prompt': entry["prompt"]}
        if entry.get("model"):
            parser[name]['model'] = entry["model"]

    def _safe_write(target_path):
        try:
            target_path.parent.mkdir(parents=True, exist_ok=True)
            with open(target_path, 'w', encoding='utf-8') as f:
                parser.write(f)
        except Exception:
            # Não quebrar o fluxo do addon se salvar em um dos destinos falhar.
            pass

    _safe_write(PROMPTS_INI_PATH)
    _safe_write(USER_PROMPTS_PATH)

prompts_collection = load_prompts()

# Especificação da configuração
confspec = {
    "api_key": "string(default='')",
    "selected_prompt": f"string(default='{DEFAULT_PROMPT_NAME}')",
    "model": "string(default='gpt-5-mini')",
    "image_prompt": "string(default='Descreva a imagem, extraia qualquer texto visível com fidelidade e entregue o resultado final em português, pronto para substituir a área de transferência.')",
    "image_model": "string(default='gpt-4o')",
}
config.conf.spec["clipboardProcessor"] = confspec
MODEL_CHOICES = ["gpt-5.2" ,"gpt-5-mini", "gpt-4o", "gpt-4.1", "gpt-4.1-mini"]
IMAGE_EXTENSIONS = {
    ".bmp", ".dib", ".gif", ".heic", ".heif", ".ico", ".jfif", ".jpeg",
    ".jpg", ".png", ".tif", ".tiff", ".webp"
}
IMAGE_MIME_PREFIXES = ("image/",)
AUDIO_EXTENSIONS = {
    ".aac", ".aiff", ".amr", ".flac", ".m4a", ".mid", ".midi", ".mp3",
    ".ogg", ".opus", ".wav", ".wma"
}
AUDIO_MIME_PREFIXES = ("audio/",)
MAX_WEB_DOWNLOAD_BYTES = 2 * 1024 * 1024
MAX_TEXT_CHARS_FOR_MODEL = 12000
DEFAULT_AUDIO_TRANSCRIPTION_MODEL = "gpt-4o-transcribe-diarize"


class _HTMLTextExtractor(HTMLParser):
    def __init__(self):
        super(_HTMLTextExtractor, self).__init__()
        self._ignored_tags = 0
        self._in_title = False
        self.title = ""
        self._chunks = []

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "noscript"):
            self._ignored_tags += 1
        elif tag == "title":
            self._in_title = True

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript") and self._ignored_tags:
            self._ignored_tags -= 1
        elif tag == "title":
            self._in_title = False
        elif tag in ("p", "br", "div", "li", "section", "article", "h1", "h2", "h3", "h4", "h5", "h6"):
            self._chunks.append("\n")

    def handle_data(self, data):
        if self._ignored_tags:
            return
        cleaned = re.sub(r"\s+", " ", data).strip()
        if not cleaned:
            return
        if self._in_title and not self.title:
            self.title = cleaned
        self._chunks.append(cleaned)

    def get_text(self):
        return re.sub(r"\n{3,}", "\n\n", "\n".join(self._chunks)).strip()

# --- Janela para Adicionar/Editar Prompt ---
class PromptDialog(wx.Dialog):
    def __init__(self, parent, title, existing_names, name="", prompt=""):
        super(PromptDialog, self).__init__(parent, title=title, size=(500, 300))
        self.existing_names = existing_names
        self.original_name = name

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        nameSizer = wx.BoxSizer(wx.HORIZONTAL)
        nameLabel = wx.StaticText(self, label="Nome do Prompt:")
        nameSizer.Add(nameLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=5)
        self.nameCtrl = wx.TextCtrl(self, value=name)
        nameSizer.Add(self.nameCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        mainSizer.Add(nameSizer, flag=wx.EXPAND)

        promptLabel = wx.StaticText(self, label="Conteúdo do Prompt:")
        mainSizer.Add(promptLabel, flag=wx.ALL, border=5)
        self.promptCtrl = wx.TextCtrl(self, value=prompt, style=wx.TE_MULTILINE)
        mainSizer.Add(self.promptCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)

        buttonSizer = wx.StdDialogButtonSizer()
        saveButton = wx.Button(self, wx.ID_OK, label="Salvar")
        cancelButton = wx.Button(self, wx.ID_CANCEL, label="Cancelar")
        buttonSizer.AddButton(saveButton)
        buttonSizer.AddButton(cancelButton)
        buttonSizer.Realize()
        mainSizer.Add(buttonSizer, flag=wx.ALIGN_CENTER | wx.ALL, border=5)

        self.SetSizer(mainSizer)
        self.nameCtrl.SetFocus()
        self.Bind(wx.EVT_BUTTON, self.on_save, id=wx.ID_OK)

    def on_save(self, event):
        name = self.nameCtrl.GetValue().strip()
        prompt = self.promptCtrl.GetValue().strip()

        if not name or not prompt:
            wx.MessageBox("O nome e o conteúdo do prompt não podem estar vazios.", "Erro de Validação", wx.OK | wx.ICON_ERROR)
            return

        invalid_found = [c for c in INVALID_PROMPT_NAME_CHARS if c in name]
        if invalid_found:
            bad_list = "', '".join(invalid_found)
            wx.MessageBox(f"O nome do prompt não pode conter: '{bad_list}'. Remova esses caracteres e tente novamente.", "Erro de Validação", wx.OK | wx.ICON_ERROR)
            return

        if name != self.original_name and name in self.existing_names:
            wx.MessageBox(f"O nome de prompt '{name}' já existe.", "Erro de Validação", wx.OK | wx.ICON_ERROR)
            return
        
        self.EndModal(wx.ID_OK)

    def get_values(self):
        return self.nameCtrl.GetValue().strip(), self.promptCtrl.GetValue().strip()

# --- Diálogo de Prompt Rápido ---
class QuickPromptDialog(wx.Dialog):
    def __init__(self, parent, default_model):
        super(QuickPromptDialog, self).__init__(parent, title="Execução de Prompt Rápido", size=(500, 250))
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        promptLabel = wx.StaticText(self, label="Escreva seu prompt:")
        mainSizer.Add(promptLabel, flag=wx.ALL, border=5)
        self.promptCtrl = wx.TextCtrl(self, style=wx.TE_MULTILINE)
        # Acessibilidade: Associar o rótulo ao campo
        self.promptCtrl.SetLabel("Campo de edição do prompt")
        mainSizer.Add(self.promptCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        
        modelSizer = wx.BoxSizer(wx.HORIZONTAL)
        modelLabel = wx.StaticText(self, label="Selecione ou digite o Modelo:")
        modelSizer.Add(modelLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=5)
        self.modelCtrl = wx.ComboBox(
            self,
            value=default_model,
            choices=MODEL_CHOICES,
            style=wx.CB_DROPDOWN # Permitir edição manual
        )
        self.modelCtrl.SetLabel("Seleção de modelo de IA")
        modelSizer.Add(self.modelCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        mainSizer.Add(modelSizer, flag=wx.EXPAND)
        
        buttonSizer = wx.StdDialogButtonSizer()
        processButton = wx.Button(self, wx.ID_OK, label="Processar")
        cancelButton = wx.Button(self, wx.ID_CANCEL, label="Cancelar")
        buttonSizer.AddButton(processButton)
        buttonSizer.AddButton(cancelButton)
        buttonSizer.Realize()
        mainSizer.Add(buttonSizer, flag=wx.ALIGN_CENTER | wx.ALL, border=5)
        
        self.SetSizer(mainSizer)
        self.promptCtrl.SetFocus()
        
        self.Bind(wx.EVT_BUTTON, self.on_process, id=wx.ID_OK)

    def on_process(self, event):
        if not self.promptCtrl.GetValue().strip():
            wx.MessageBox("O prompt não pode estar vazio.", "Erro", wx.OK | wx.ICON_ERROR)
            return
        self.EndModal(wx.ID_OK)

    def get_values(self):
        return self.promptCtrl.GetValue().strip(), self.modelCtrl.GetValue().strip()

# --- Painel de Configurações Principal ---
class SettingsPanel(settingsDialogs.SettingsPanel):
    title = "Processador de Clipboard com IA"

    def makeSettings(self, settingsSizer):
        self.edited_prompts = {name: _normalize_prompt_entry(value) for name, value in prompts_collection.items()}

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        # Configurações Gerais
        generalBox = wx.StaticBoxSizer(wx.VERTICAL, self, label="Configurações Gerais")
        apiKeySizer = wx.BoxSizer(wx.HORIZONTAL)
        apiKeyLabel = wx.StaticText(generalBox.GetStaticBox(), label="Chave da API:")
        apiKeySizer.Add(apiKeyLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=5)
        self.apiKeyCtrl = wx.TextCtrl(generalBox.GetStaticBox(), value=config.conf["clipboardProcessor"]["api_key"], style=wx.TE_PASSWORD)
        apiKeySizer.Add(self.apiKeyCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        generalBox.Add(apiKeySizer, flag=wx.EXPAND)

        modelSizer = wx.BoxSizer(wx.HORIZONTAL)
        modelLabel = wx.StaticText(generalBox.GetStaticBox(), label="Modelo para Texto e URL (selecione ou digite):")
        modelSizer.Add(modelLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=5)
        self.modelCtrl = wx.ComboBox(
            generalBox.GetStaticBox(),
            value=config.conf["clipboardProcessor"]["model"],
            choices=MODEL_CHOICES,
            style=wx.CB_DROPDOWN
        )
        modelSizer.Add(self.modelCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        generalBox.Add(modelSizer, flag=wx.EXPAND)

        defaultPromptSizer = wx.BoxSizer(wx.HORIZONTAL)
        defaultPromptLabel = wx.StaticText(generalBox.GetStaticBox(), label="Prompt Padrão para Texto e URL:")
        defaultPromptSizer.Add(defaultPromptLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=5)
        self.defaultPromptCtrl = wx.ComboBox(generalBox.GetStaticBox(), value=config.conf["clipboardProcessor"]["selected_prompt"], choices=list(self.edited_prompts.keys()), style=wx.CB_READONLY)
        defaultPromptSizer.Add(self.defaultPromptCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        generalBox.Add(defaultPromptSizer, flag=wx.EXPAND)

        generalInfoLabel = wx.StaticText(
            generalBox.GetStaticBox(),
            label="Texto e URL usam os prompts abaixo. Áudio não usa prompt: ele gera transcrição com falantes e tempo."
        )
        generalInfoLabel.Wrap(560)
        generalBox.Add(generalInfoLabel, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=5)

        imageModelSizer = wx.BoxSizer(wx.HORIZONTAL)
        imageModelLabel = wx.StaticText(generalBox.GetStaticBox(), label="Modelo Exclusivo para Imagem:")
        imageModelSizer.Add(imageModelLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=5)
        self.imageModelCtrl = wx.ComboBox(
            generalBox.GetStaticBox(),
            value=config.conf["clipboardProcessor"]["image_model"],
            choices=MODEL_CHOICES,
            style=wx.CB_DROPDOWN
        )
        imageModelSizer.Add(self.imageModelCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        generalBox.Add(imageModelSizer, flag=wx.EXPAND)

        imagePromptLabel = wx.StaticText(generalBox.GetStaticBox(), label="Prompt Exclusivo para Imagem/Bitmap:")
        generalBox.Add(imagePromptLabel, flag=wx.LEFT | wx.RIGHT | wx.TOP, border=5)
        self.imagePromptCtrl = wx.TextCtrl(
            generalBox.GetStaticBox(),
            value=config.conf["clipboardProcessor"]["image_prompt"],
            style=wx.TE_MULTILINE
        )
        generalBox.Add(self.imagePromptCtrl, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)

        imageInfoLabel = wx.StaticText(
            generalBox.GetStaticBox(),
            label="Este prompt é usado somente quando o clipboard contém um arquivo de imagem ou bitmap copiado."
        )
        imageInfoLabel.Wrap(560)
        generalBox.Add(imageInfoLabel, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=5)
        mainSizer.Add(generalBox, flag=wx.EXPAND | wx.ALL, border=5)

        # Gerenciador de Prompts
        promptsBox = wx.StaticBoxSizer(wx.VERTICAL, self, label="Prompts de Texto e URL")
        managerSizer = wx.BoxSizer(wx.HORIZONTAL)

        promptsInfoLabel = wx.StaticText(
            promptsBox.GetStaticBox(),
            label="Esses prompts aparecem para seleção manual em texto e URL. Eles não são usados no fluxo de áudio nem no fluxo exclusivo de imagem."
        )
        promptsInfoLabel.Wrap(560)
        promptsBox.Add(promptsInfoLabel, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=5)
        
        self.promptListCtrl = wx.ListBox(promptsBox.GetStaticBox(), choices=list(self.edited_prompts.keys()))
        managerSizer.Add(self.promptListCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)

        editorSizer = wx.BoxSizer(wx.VERTICAL)
        self.promptContentCtrl = wx.TextCtrl(promptsBox.GetStaticBox(), style=wx.TE_MULTILINE)
        editorSizer.Add(self.promptContentCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)

        modelRow = wx.BoxSizer(wx.HORIZONTAL)
        modelLabel = wx.StaticText(promptsBox.GetStaticBox(), label="Modelo deste prompt (opcional):")
        modelRow.Add(modelLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=5)
        self.promptModelCtrl = wx.ComboBox(
            promptsBox.GetStaticBox(),
            choices=[""] + MODEL_CHOICES,
            style=wx.CB_DROPDOWN
        )
        modelRow.Add(self.promptModelCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        editorSizer.Add(modelRow, flag=wx.EXPAND)
        
        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.saveEditBtn = wx.Button(promptsBox.GetStaticBox(), label="Salvar Edição")
        self.deleteBtn = wx.Button(promptsBox.GetStaticBox(), label="Excluir")
        self.addBtn = wx.Button(promptsBox.GetStaticBox(), label="Adicionar...")
        buttonSizer.Add(self.saveEditBtn, flag=wx.ALL, border=5)
        buttonSizer.Add(self.deleteBtn, flag=wx.ALL, border=5)
        buttonSizer.Add(self.addBtn, flag=wx.ALL, border=5)
        editorSizer.Add(buttonSizer, flag=wx.ALIGN_CENTER)
        
        managerSizer.Add(editorSizer, proportion=2, flag=wx.EXPAND)
        promptsBox.Add(managerSizer, proportion=1, flag=wx.EXPAND)
        mainSizer.Add(promptsBox, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)

        settingsSizer.Add(mainSizer, proportion=1, flag=wx.EXPAND)

        # Bind Events
        self.promptListCtrl.Bind(wx.EVT_LISTBOX, self.on_prompt_selected)
        self.saveEditBtn.Bind(wx.EVT_BUTTON, self.on_save_edit)
        self.deleteBtn.Bind(wx.EVT_BUTTON, self.on_delete)
        self.addBtn.Bind(wx.EVT_BUTTON, self.on_add)

        self.update_controls_state()

    def on_prompt_selected(self, event):
        sel_name = self.promptListCtrl.GetStringSelection()
        if sel_name:
            entry = _normalize_prompt_entry(self.edited_prompts.get(sel_name, {}))
            self.promptContentCtrl.SetValue(entry.get("prompt", ""))
            self.promptModelCtrl.SetValue(entry.get("model") or "")
        self.update_controls_state()

    def on_save_edit(self, event):
        sel_name = self.promptListCtrl.GetStringSelection()
        if not sel_name:
            return
        new_content = self.promptContentCtrl.GetValue()
        selected_model = self.promptModelCtrl.GetValue().strip() or None
        self.edited_prompts[sel_name] = {"prompt": new_content, "model": selected_model}
        ui.message("Edição salva na memória. Clique em OK ou Aplicar para salvar no arquivo.")

    def on_delete(self, event):
        sel_name = self.promptListCtrl.GetStringSelection()
        if not sel_name:
            return
        
        if len(self.edited_prompts) <= 1:
            wx.MessageBox("Você não pode excluir o último prompt.", "Ação Inválida", wx.OK | wx.ICON_WARNING)
            return

        del self.edited_prompts[sel_name]
        self.refresh_prompt_lists()
        self.promptContentCtrl.Clear()
        self.update_controls_state()

    def on_add(self, event):
        dialog = PromptDialog(self, "Adicionar Novo Prompt", self.edited_prompts.keys())
        if dialog.ShowModal() == wx.ID_OK:
            name, prompt = dialog.get_values()
            self.edited_prompts[name] = {"prompt": prompt, "model": None}
            self.refresh_prompt_lists(new_selection=name)
        dialog.Destroy()

    def refresh_prompt_lists(self, new_selection=None):
        prompt_names = list(self.edited_prompts.keys())
        
        # Salva a seleção atual para tentar restaurá-la
        current_default = self.defaultPromptCtrl.GetValue()
        
        # Atualiza a lista principal
        self.promptListCtrl.Set(prompt_names)
        if new_selection:
            self.promptListCtrl.SetStringSelection(new_selection)
            self.on_prompt_selected(None) # Atualiza o campo de texto
        
        # Atualiza a lista de prompts padrão
        self.defaultPromptCtrl.Set(prompt_names)
        if current_default in prompt_names:
            self.defaultPromptCtrl.SetValue(current_default)
        elif prompt_names:
            self.defaultPromptCtrl.SetValue(prompt_names[0])
        else:
            self.defaultPromptCtrl.SetValue("")

    def update_controls_state(self):
        has_selection = self.promptListCtrl.GetSelection() != wx.NOT_FOUND
        self.saveEditBtn.Enable(has_selection)
        self.deleteBtn.Enable(has_selection)
        self.promptContentCtrl.Enable(has_selection)
        self.promptModelCtrl.Enable(has_selection)

    def onSave(self):
        global prompts_collection
        config.conf["clipboardProcessor"]["api_key"] = self.apiKeyCtrl.GetValue()
        config.conf["clipboardProcessor"]["model"] = self.modelCtrl.GetValue().strip()
        config.conf["clipboardProcessor"]["selected_prompt"] = self.defaultPromptCtrl.GetValue()
        config.conf["clipboardProcessor"]["image_model"] = self.imageModelCtrl.GetValue().strip()
        config.conf["clipboardProcessor"]["image_prompt"] = self.imagePromptCtrl.GetValue().strip()
        
        save_prompts(self.edited_prompts)
        prompts_collection = copy.deepcopy(self.edited_prompts)

# --- Lógica do Plugin Global ---
class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    
    script_category = "Processador de Clipboard com IA"

    def __init__(self, *args, **kwargs):
        super(GlobalPlugin, self).__init__(*args, **kwargs)
        self.session = requests.Session() if requests else None
        if not config.conf["clipboardProcessor"]["api_key"]:
            env_key = os.environ.get('OPENAI_API_KEY')
            if env_key:
                config.conf["clipboardProcessor"]["api_key"] = env_key
        settingsDialogs.NVDASettingsDialog.categoryClasses.append(SettingsPanel)

    def terminate(self, *args, **kwargs):
        try:
            settingsDialogs.NVDASettingsDialog.categoryClasses.remove(SettingsPanel)
        except ValueError:
            pass
        if self.session:
            self.session.close()
        super(GlobalPlugin, self).terminate(*args, **kwargs)

    def _get_prompt_entry(self, selected_prompt_name):
        prompt_entry = _normalize_prompt_entry(prompts_collection.get(selected_prompt_name))
        system_prompt = prompt_entry.get("prompt")
        if not system_prompt:
            raise ValueError(f"Prompt '{selected_prompt_name}' não encontrado.")
        return prompt_entry

    def _ensure_api_requirements(self):
        api_key = config.conf["clipboardProcessor"]["api_key"]

        if not requests or not self.session:
            raise RuntimeError("Erro: A biblioteca 'requests' não foi encontrada.")

        if not api_key:
            raise RuntimeError("Por favor, configure sua chave da API da OpenAI.")

        return api_key

    def _handle_api_error(self, response):
        """Tratamento granular de erros da API OpenAI."""
        if response.status_code == 401:
            return "Chave de API inválida ou expirada. Verifique as configurações."
        if response.status_code == 429:
            return "Limite de cota atingido ou excesso de requisições. Tente novamente mais tarde."
        if response.status_code >= 500:
            return f"Erro no servidor da OpenAI ({response.status_code}). Tente novamente em instantes."
        try:
            err_data = response.json()
            msg = err_data.get("error", {}).get("message", "Erro desconhecido na API.")
            return f"Erro da API: {msg}"
        except Exception:
            return f"Erro na requisição: Código {response.status_code}"

    def _extract_chat_result_text(self, json_response):
        result_text = json_response["choices"][0]["message"]["content"]
        if isinstance(result_text, list):
            parts = []
            for item in result_text:
                if isinstance(item, dict) and item.get("type") == "text":
                    parts.append(item.get("text", ""))
            result_text = "\n".join(part for part in parts if part)
        return result_text

    def _extract_responses_result_text(self, json_response):
        output_items = json_response.get("output", [])
        text_parts = []
        for item in output_items:
            for content_item in item.get("content", []):
                if content_item.get("type") == "output_text":
                    text_parts.append(content_item.get("text", ""))
        if not text_parts and json_response.get("output_text"):
            text_parts.append(json_response.get("output_text", ""))
        return "\n".join(part for part in text_parts if part).strip()

    def _post_chat_completions(self, system_prompt, user_text, model):
        api_key = self._ensure_api_requirements()
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_text},
            ],
        }
        response = self.session.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=90
        )
        if not response.ok:
            raise RuntimeError(self._handle_api_error(response))
        return self._sanitize_markdown_output(self._extract_chat_result_text(response.json()))

    def _post_responses_api(self, model, instructions, input_content):
        api_key = self._ensure_api_requirements()
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        payload = {
            "model": model,
            "input": [
                {
                    "role": "user",
                    "content": input_content,
                }
            ],
            "instructions": instructions,
        }
        response = self.session.post(
            "https://api.openai.com/v1/responses",
            headers=headers,
            json=payload,
            timeout=120
        )
        if not response.ok:
            raise RuntimeError(self._handle_api_error(response))
        return self._sanitize_markdown_output(self._extract_responses_result_text(response.json()))

    def _process_text_with_prompt(self, text_to_process, selected_prompt_name):
        prompt_entry = self._get_prompt_entry(selected_prompt_name)
        prompt_model = prompt_entry.get("model") or config.conf["clipboardProcessor"]["model"]
        return self._post_chat_completions(
            prompt_entry.get("prompt"),
            text_to_process,
            prompt_model
        )

    def _worker_thread(self, text_to_process, selected_prompt_name):
        try:
            clean_text = self._process_text_with_prompt(text_to_process, selected_prompt_name)
            wx.CallAfter(self._update_clipboard, clean_text)
        except ValueError as e:
            wx.CallAfter(ui.message, f"Erro: {e}")
        except RuntimeError as e:
            wx.CallAfter(ui.message, str(e))
        except requests.exceptions.Timeout:
            wx.CallAfter(ui.message, "Tempo limite ao chamar a API. Tente novamente.")
        except requests.exceptions.RequestException as e:
            wx.CallAfter(ui.message, f"Erro de conexão: {e}")
        except (KeyError, IndexError):
            wx.CallAfter(ui.message, "Erro: Resposta inesperada da API.")
        except Exception as e:
            wx.CallAfter(ui.message, f"Ocorreu um erro inesperado: {e}")

    def _bitmap_to_png_data_url(self, bitmap):
        fd, temp_path = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        try:
            image = bitmap.ConvertToImage()
            if not image.IsOk():
                raise RuntimeError("Não foi possível converter a imagem da área de transferência.")
            if not image.SaveFile(temp_path, wx.BITMAP_TYPE_PNG):
                raise RuntimeError("Não foi possível serializar a imagem da área de transferência.")
            with open(temp_path, "rb") as temp_file:
                encoded = base64.b64encode(temp_file.read()).decode("ascii")
            return f"data:image/png;base64,{encoded}"
        finally:
            try:
                os.remove(temp_path)
            except OSError:
                pass

    def _image_file_to_data_url(self, file_path):
        path_obj = Path(file_path)
        mime_type, _ = mimetypes.guess_type(str(path_obj))
        mime_type = mime_type or "application/octet-stream"
        with open(path_obj, "rb") as image_file:
            encoded = base64.b64encode(image_file.read()).decode("ascii")
        return f"data:{mime_type};base64,{encoded}"

    def _process_image_with_prompt(self, payload):
        if payload.get("kind") == "blob_url_image":
            raise RuntimeError(
                "URL blob detectada, mas esse formato não é acessível fora do aplicativo de origem. "
                "Copie a imagem em si ou salve-a como arquivo."
            )

        prompt_model = config.conf["clipboardProcessor"]["image_model"] or config.conf["clipboardProcessor"]["model"]
        image_prompt = config.conf["clipboardProcessor"]["image_prompt"].strip()
        if not image_prompt:
            raise RuntimeError("Configure o prompt exclusivo de imagem nas configurações do addon.")

        if payload.get("kind") == "image_bitmap":
            image_url = self._bitmap_to_png_data_url(payload["bitmap"])
            image_label = "imagem copiada da área de transferência"
        else:
            image_url = self._image_file_to_data_url(payload["path"])
            image_label = payload.get("display") or "arquivo de imagem"

        input_content = [
            {
                "type": "input_text",
                "text": (
                    "Analise a imagem fornecida e responda em português. "
                    "Se houver texto legível, transcreva-o fielmente antes de aplicar a transformação solicitada."
                ),
            },
            {
                "type": "input_image",
                "image_url": image_url,
            },
        ]
        instructions = (
            f"{image_prompt}\n\n"
            "A entrada do usuário é uma imagem. Descreva apenas o que estiver presente nela "
            "e produza uma saída final pronta para substituir a área de transferência."
        )
        return self._post_responses_api(prompt_model, instructions, input_content), image_label

    def _format_seconds_to_timestamp(self, seconds_value):
        try:
            total_milliseconds = int(round(float(seconds_value) * 1000))
        except (TypeError, ValueError):
            return "00:00:00.000"
        hours = total_milliseconds // 3600000
        remaining = total_milliseconds % 3600000
        minutes = remaining // 60000
        remaining %= 60000
        seconds = remaining // 1000
        milliseconds = remaining % 1000
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}.{milliseconds:03d}"

    def _build_diarized_transcript(self, transcription_json):
        segments = transcription_json.get("segments") or []
        if not segments:
            text_output = (transcription_json.get("text") or "").strip()
            if not text_output:
                raise RuntimeError("A transcrição do áudio retornou vazia.")
            return text_output

        lines = []
        for segment in segments:
            speaker = segment.get("speaker") or "Falante"
            start_time = self._format_seconds_to_timestamp(segment.get("start"))
            end_time = self._format_seconds_to_timestamp(segment.get("end"))
            text = (segment.get("text") or "").strip()
            if not text:
                continue
            lines.append(f"[{start_time} --> {end_time}] {speaker}: {text}")

        if not lines:
            raise RuntimeError("A diarização do áudio não retornou segmentos utilizáveis.")
        return "\n".join(lines)

    def _transcribe_audio_file(self, audio_path):
        api_key = self._ensure_api_requirements()
        headers = {"Authorization": f"Bearer {api_key}"}
        with open(audio_path, "rb") as audio_file:
            files = {"file": (Path(audio_path).name, audio_file)}
            data = {
                "model": DEFAULT_AUDIO_TRANSCRIPTION_MODEL,
                "response_format": "diarized_json",
            }
            response = self.session.post(
                "https://api.openai.com/v1/audio/transcriptions",
                headers=headers,
                files=files,
                data=data,
                timeout=180
            )
        if not response.ok:
            raise RuntimeError(self._handle_api_error(response))
        return response.json()

    def _process_audio_transcription(self, payload):
        audio_path = payload.get("path")
        if not audio_path:
            raise RuntimeError("Nenhum arquivo de áudio válido foi encontrado.")
        transcription_json = self._transcribe_audio_file(audio_path)
        final_text = self._build_diarized_transcript(transcription_json)
        return final_text, payload.get("display") or "áudio"

    def _download_web_resource(self, url):
        response = self.session.get(
            url,
            headers={"User-Agent": "clipboardProcessor/1.1"},
            stream=True,
            timeout=30
        )
        if not response.ok:
            raise RuntimeError(f"Erro ao acessar a URL: Código {response.status_code}")
        content_type = (response.headers.get("Content-Type") or "").split(";")[0].strip().lower()
        chunks = []
        total_bytes = 0
        for chunk in response.iter_content(chunk_size=8192):
            if not chunk:
                continue
            total_bytes += len(chunk)
            if total_bytes > MAX_WEB_DOWNLOAD_BYTES:
                break
            chunks.append(chunk)
        raw_bytes = b"".join(chunks)
        return response.url, content_type, raw_bytes

    def _extract_text_from_web_response(self, final_url, content_type, raw_bytes):
        if not raw_bytes:
            raise RuntimeError("A URL não retornou conteúdo legível.")

        if content_type.startswith("text/html") or content_type == "":
            decoded = raw_bytes.decode("utf-8", errors="replace")
            extractor = _HTMLTextExtractor()
            extractor.feed(decoded)
            extracted_text = extractor.get_text()
            if not extracted_text:
                raise RuntimeError("Não foi possível extrair texto relevante da página.")
            page_title = extractor.title or final_url
            return page_title, extracted_text

        if content_type.startswith("text/") or content_type in ("application/json", "application/xml", "text/xml"):
            decoded = raw_bytes.decode("utf-8", errors="replace").strip()
            if not decoded:
                raise RuntimeError("O conteúdo textual da URL está vazio.")
            return final_url, decoded

        raise RuntimeError(f"Tipo de conteúdo da URL não suportado: {content_type or 'desconhecido'}.")

    def _truncate_for_model(self, text):
        clean_text = text.strip()
        if len(clean_text) <= MAX_TEXT_CHARS_FOR_MODEL:
            return clean_text
        return clean_text[:MAX_TEXT_CHARS_FOR_MODEL] + "\n\n[Conteúdo truncado para caber no limite de processamento.]"

    def _process_web_url_with_prompt(self, payload, selected_prompt_name):
        url = payload.get("text")
        if not url:
            raise RuntimeError("Nenhuma URL válida foi encontrada.")
        final_url, content_type, raw_bytes = self._download_web_resource(url)
        page_title, extracted_text = self._extract_text_from_web_response(final_url, content_type, raw_bytes)
        prepared_text = self._truncate_for_model(
            f"URL: {final_url}\n"
            f"Título: {page_title}\n"
            f"Tipo de conteúdo: {content_type or 'desconhecido'}\n\n"
            f"Conteúdo extraído:\n{extracted_text}"
        )
        final_text = self._process_text_with_prompt(prepared_text, selected_prompt_name)
        return final_text, payload.get("display") or page_title

    def _classify_file_path(self, file_path):
        normalized_path = Path(file_path).expanduser()
        suffix = normalized_path.suffix.lower()
        mime_type, _ = mimetypes.guess_type(str(normalized_path))

        if suffix in IMAGE_EXTENSIONS or (mime_type and mime_type.startswith(IMAGE_MIME_PREFIXES)):
            return "image_file"
        if suffix in AUDIO_EXTENSIONS or (mime_type and mime_type.startswith(AUDIO_MIME_PREFIXES)):
            return "audio_file"
        return "file"

    def _looks_like_single_local_path(self, text):
        candidate = text.strip().strip('"')
        if not candidate or "\n" in candidate or "\r" in candidate:
            return None
        if candidate.startswith("\\\\") or re.match(r"^[a-zA-Z]:[\\/]", candidate):
            candidate_path = Path(candidate)
            if candidate_path.exists():
                return str(candidate_path)
        return None

    def _classify_text_payload(self, text):
        stripped_text = text.strip()
        lowered_text = stripped_text.lower()

        if lowered_text.startswith("blob:"):
            return {
                "kind": "blob_url_image",
                "source": "text",
                "text": stripped_text,
                "display": "URL blob",
            }

        if lowered_text.startswith(("http://", "https://")):
            parsed_url = urlparse(stripped_text)
            return {
                "kind": "web_url",
                "source": "text",
                "text": stripped_text,
                "display": parsed_url.netloc or "URL web",
            }

        local_path = self._looks_like_single_local_path(stripped_text)
        if local_path:
            return {
                "kind": self._classify_file_path(local_path),
                "source": "text",
                "path": local_path,
                "text": stripped_text,
                "display": Path(local_path).name,
            }

        return {
            "kind": "text",
            "source": "text",
            "text": text,
            "display": "texto",
        }

    def _read_clipboard_payload(self):
        payload = {
            "kind": "empty",
            "source": None,
            "display": "",
        }
        clipboard_opened = False

        try:
            clipboard_opened = wx.TheClipboard.Open()
            if not clipboard_opened:
                raise RuntimeError("Não foi possível abrir a área de transferência.")

            if wx.TheClipboard.IsSupported(wx.DataFormat(wx.DF_FILENAME)):
                file_data = wx.FileDataObject()
                if wx.TheClipboard.GetData(file_data):
                    filenames = list(file_data.GetFilenames() or [])
                    if filenames:
                        primary_path = filenames[0]
                        payload = {
                            "kind": self._classify_file_path(primary_path),
                            "source": "file_drop",
                            "path": primary_path,
                            "paths": filenames,
                            "display": Path(primary_path).name,
                        }
                        if len(filenames) > 1:
                            payload["kind"] = "multiple_files"
                        return payload

            if wx.TheClipboard.IsSupported(wx.DataFormat(wx.DF_BITMAP)):
                bitmap_data = wx.BitmapDataObject()
                if wx.TheClipboard.GetData(bitmap_data):
                    bitmap = bitmap_data.GetBitmap()
                    if bitmap.IsOk():
                        return {
                            "kind": "image_bitmap",
                            "source": "bitmap",
                            "bitmap": bitmap,
                            "display": "imagem",
                        }

            if wx.TheClipboard.IsSupported(wx.DataFormat(wx.DF_TEXT)):
                text_data = wx.TextDataObject()
                if wx.TheClipboard.GetData(text_data):
                    text = text_data.GetText()
                    if text and text.strip():
                        return self._classify_text_payload(text)

            return payload
        finally:
            if clipboard_opened:
                wx.TheClipboard.Close()

    def _handle_text_content(self, payload):
        text_to_process = payload.get("text", "")
        if not text_to_process or not text_to_process.strip():
            ui.message("Nenhum texto para processar.")
            return
        wx.CallAfter(
            self._show_prompt_selection_menu,
            lambda selected_prompt_name: self._start_text_processing(text_to_process, selected_prompt_name)
        )

    def _handle_image_content(self, payload):
        self._start_image_processing(payload)

    def _handle_audio_content(self, payload):
        self._start_audio_processing(payload)

    def _handle_web_url_content(self, payload):
        wx.CallAfter(
            self._show_prompt_selection_menu,
            lambda selected_prompt_name: self._start_web_url_processing(payload, selected_prompt_name)
        )

    def _dispatch_clipboard_payload(self, payload):
        kind = payload.get("kind")

        if kind == "text":
            self._handle_text_content(payload)
            return

        if kind in ("image_file", "image_bitmap", "blob_url_image"):
            self._handle_image_content(payload)
            return

        if kind == "audio_file":
            self._handle_audio_content(payload)
            return

        if kind == "web_url":
            self._handle_web_url_content(payload)
            return

        if kind == "multiple_files":
            ui.message("Foram detectados múltiplos arquivos na área de transferência. Copie apenas um item por vez.")
            return

        if kind == "file":
            ui.message("O arquivo copiado não é um formato de imagem ou áudio suportado.")
            return

        ui.message("Nenhum conteúdo suportado foi identificado na área de transferência.")

    def _sanitize_markdown_output(self, text):
        """Remove marcadores comuns de Markdown, preservando quebras de linha e t¡tulos com #."""
        # Remove blocos de c¢digo delimitadores, mantendo o conte£do interno.
        text = re.sub(r"```(.*?)```", r"\1", text, flags=re.DOTALL)
        # Remover marcadores de lista no in¡cio da linha: -, *, +, • ou n£meros seguidos de . ou ).
        text = re.sub(r"(?m)^\s*([-*+•]|\d+[.)])\s+", "", text)
        # Remover blockquote '>' no in¡cio da linha.
        text = re.sub(r"(?m)^\s*>\s?", "", text)
        # Remover ˆnfases em negrito/it lico e inline code, preservando o texto.
        text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
        text = re.sub(r"__([^_]+)__", r"\1", text)
        text = re.sub(r"\*([^*]+)\*", r"\1", text)
        text = re.sub(r"_([^_]+)_", r"\1", text)
        text = re.sub(r"`([^`]*)`", r"\1", text)
        return text

    def _update_clipboard(self, text):
        try:
            if wx.TheClipboard.Open():
                wx.TheClipboard.SetData(wx.TextDataObject(text))
                wx.TheClipboard.Close()
                tones.beep(800, 100)
                ui.message(text)
        except Exception as e:
            ui.message(f"Erro ao atualizar a área de transferência: {e}")

    def _start_background_task(self, announcement, worker, *args):
        ui.message(announcement)
        tones.beep(300, 100)
        thread = threading.Thread(target=worker, args=args)
        thread.start()

    def _start_text_processing(self, text_to_process, selected_prompt_name):
        self._start_background_task(
            f"Processando com '{selected_prompt_name}'...",
            self._worker_thread,
            text_to_process,
            selected_prompt_name
        )

    def _image_worker_thread(self, payload):
        try:
            clean_text, image_label = self._process_image_with_prompt(payload)
            wx.CallAfter(self._update_clipboard, clean_text)
        except RuntimeError as e:
            wx.CallAfter(ui.message, str(e))
        except ValueError as e:
            wx.CallAfter(ui.message, f"Erro: {e}")
        except requests.exceptions.Timeout:
            wx.CallAfter(ui.message, "Tempo limite ao processar a imagem.")
        except requests.exceptions.RequestException as e:
            wx.CallAfter(ui.message, f"Erro de conexão: {e}")
        except Exception as e:
            wx.CallAfter(ui.message, f"Ocorreu um erro ao processar a imagem: {e}")

    def _audio_worker_thread(self, payload):
        try:
            clean_text, audio_label = self._process_audio_transcription(payload)
            wx.CallAfter(self._update_clipboard, clean_text)
        except RuntimeError as e:
            wx.CallAfter(ui.message, str(e))
        except ValueError as e:
            wx.CallAfter(ui.message, f"Erro: {e}")
        except requests.exceptions.Timeout:
            wx.CallAfter(ui.message, "Tempo limite ao processar o áudio.")
        except requests.exceptions.RequestException as e:
            wx.CallAfter(ui.message, f"Erro de conexão: {e}")
        except Exception as e:
            wx.CallAfter(ui.message, f"Ocorreu um erro ao processar o áudio: {e}")

    def _web_url_worker_thread(self, payload, selected_prompt_name):
        try:
            clean_text, page_label = self._process_web_url_with_prompt(payload, selected_prompt_name)
            wx.CallAfter(self._update_clipboard, clean_text)
        except RuntimeError as e:
            wx.CallAfter(ui.message, str(e))
        except ValueError as e:
            wx.CallAfter(ui.message, f"Erro: {e}")
        except requests.exceptions.Timeout:
            wx.CallAfter(ui.message, "Tempo limite ao processar a URL.")
        except requests.exceptions.RequestException as e:
            wx.CallAfter(ui.message, f"Erro de conexão: {e}")
        except Exception as e:
            wx.CallAfter(ui.message, f"Ocorreu um erro ao processar a URL: {e}")

    def _quick_prompt_worker_thread(self, user_prompt, model):
        try:
            clean_text = self._post_chat_completions(
                "Você é um assistente útil. Responda de forma direta e concisa ao prompt do usuário.",
                user_prompt,
                model
            )
            wx.CallAfter(self._update_clipboard, clean_text)
        except Exception as e:
            wx.CallAfter(ui.message, str(e))

    def _start_image_processing(self, payload):
        self._start_background_task(
            "Processando imagem...",
            self._image_worker_thread,
            payload
        )

    def _start_audio_processing(self, payload):
        self._start_background_task(
            "Transcrevendo áudio com diarização...",
            self._audio_worker_thread,
            payload
        )

    def _start_web_url_processing(self, payload, selected_prompt_name):
        self._start_background_task(
            f"Processando URL com '{selected_prompt_name}'...",
            self._web_url_worker_thread,
            payload,
            selected_prompt_name
        )

    def _show_prompt_selection_menu(self, start_processing_callback):

        prompt_names = list(prompts_collection.keys())
        default_prompt = config.conf["clipboardProcessor"]["selected_prompt"]
        
        try:
            default_index = prompt_names.index(default_prompt)
        except ValueError:
            default_index = 0

        parent = wx.GetApp().GetTopWindow()
        dialog = wx.SingleChoiceDialog(parent, "Selecione o prompt para usar:", "Processar com IA", prompt_names)
        dialog.SetSelection(default_index)
        
        dialog.CentreOnScreen()
        dialog.Raise()

        if dialog.ShowModal() == wx.ID_OK:
            selected_prompt_name = dialog.GetStringSelection()
            start_processing_callback(selected_prompt_name)
        
        dialog.Destroy()

    def _show_quick_prompt_dialog(self):
        default_model = config.conf["clipboardProcessor"]["model"]
        parent = wx.GetApp().GetTopWindow()
        dialog = QuickPromptDialog(parent, default_model)
        
        if dialog.ShowModal() == wx.ID_OK:
            user_prompt, selected_model = dialog.get_values()
            self._start_background_task(
                f"Processando prompt rápido com {selected_model}...",
                self._quick_prompt_worker_thread,
                user_prompt,
                selected_model
            )
        
        dialog.Destroy()

    def script_quickPrompt(self, gesture):
        wx.CallAfter(self._show_quick_prompt_dialog)

    def script_processSelection(self, gesture):
        selected_text = ""
        try:
            focus = api.getFocusObject()
            selection = getattr(focus, "selection", None)
            if selection:
                selected_text = (selection.text or "").strip()
        except Exception:
            pass

        if not selected_text:
            tones.beep(200, 100)
            ui.message("Nenhum texto selecionado.")
            return

        wx.CallAfter(
            self._show_prompt_selection_menu,
            lambda selected_prompt_name: self._start_text_processing(selected_text, selected_prompt_name)
        )

    def script_processClipboard(self, gesture):
        try:
            clipboard_payload = self._read_clipboard_payload()
        except Exception as e:
            ui.message(f"Não foi possível ler a área de transferência: {e}")
            return
        self._dispatch_clipboard_payload(clipboard_payload)

    __gestures = {
        "kb:NVDA+shift+p": "processSelection",
        "kb:control+NVDA+shift+p": "processClipboard",
        "kb:control+alt+shift+NVDA+p": "quickPrompt",
    }
