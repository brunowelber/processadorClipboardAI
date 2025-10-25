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
DEFAULT_PROMPT_NAME = "Melhorar Ortografia e Gramática"
DEFAULT_PROMPTS = {
    DEFAULT_PROMPT_NAME: "Sua tarefa é refinar o texto a seguir. Mantenha o estilo e a voz originais do autor, mas melhore a clareza, a coesão e a concisão. Corrija todos os erros de ortografia e gramática. Elimine a prolixidade e torne a escrita mais direta e polida.",
    "Traduzir para Inglês": "traduza o seguinte texto para o inglês.",
    "Resumir em Pontos-Chave": "Resuma o texto a seguir em uma lista de pontos-chave.",
    "Tornar Mais Formal": "Reescreva o texto a seguir em um tom mais formal e profissional."
}

def load_prompts():
    if not PROMPTS_INI_PATH.exists():
        save_prompts(DEFAULT_PROMPTS)
    
    prompts = {}
    config_parser = configparser.ConfigParser()
    config_parser.read(PROMPTS_INI_PATH, encoding='utf-8')
    for section in config_parser.sections():
        if 'prompt' in config_parser[section]:
            prompts[section] = config_parser[section]['prompt']
    return prompts

def save_prompts(prompts_dict):
    config_parser = configparser.ConfigParser()
    for name, text in prompts_dict.items():
        config_parser[name] = {'prompt': text}
    with open(PROMPTS_INI_PATH, 'w', encoding='utf-8') as f:
        config_parser.write(f)

prompts_collection = load_prompts()

# Especificação da configuração
confspec = {
    "api_key": "string(default='')",
    "selected_prompt": f"string(default='{DEFAULT_PROMPT_NAME}')",
    "model": "string(default='gpt-4o')",
}
config.conf.spec["clipboardProcessor"] = confspec

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

        if name != self.original_name and name in self.existing_names:
            wx.MessageBox(f"O nome de prompt '{name}' já existe.", "Erro de Validação", wx.OK | wx.ICON_ERROR)
            return
        
        self.EndModal(wx.ID_OK)

    def get_values(self):
        return self.nameCtrl.GetValue().strip(), self.promptCtrl.GetValue().strip()

# --- Painel de Configurações Principal ---
class SettingsPanel(settingsDialogs.SettingsPanel):
    title = "Processador de Clipboard com IA"

    def makeSettings(self, settingsSizer):
        self.edited_prompts = copy.deepcopy(prompts_collection)

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        # Configurações Gerais
        generalBox = wx.StaticBoxSizer(wx.VERTICAL, self, label="Configurações Gerais")
        apiKeySizer = wx.BoxSizer(wx.HORIZONTAL)
        apiKeyLabel = wx.StaticText(generalBox.GetStaticBox(), label="Chave da API:")
        apiKeySizer.Add(apiKeyLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=5)
        self.apiKeyCtrl = wx.TextCtrl(generalBox.GetStaticBox(), value=config.conf["clipboardProcessor"]["api_key"])
        apiKeySizer.Add(self.apiKeyCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        generalBox.Add(apiKeySizer, flag=wx.EXPAND)

        modelSizer = wx.BoxSizer(wx.HORIZONTAL)
        modelLabel = wx.StaticText(generalBox.GetStaticBox(), label="Modelo:")
        modelSizer.Add(modelLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=5)
        self.modelCtrl = wx.ComboBox(generalBox.GetStaticBox(), value=config.conf["clipboardProcessor"]["model"], choices=['gpt-4o', 'gpt-4-turbo', 'gpt-3.5-turbo'], style=wx.CB_READONLY)
        modelSizer.Add(self.modelCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        generalBox.Add(modelSizer, flag=wx.EXPAND)

        defaultPromptSizer = wx.BoxSizer(wx.HORIZONTAL)
        defaultPromptLabel = wx.StaticText(generalBox.GetStaticBox(), label="Prompt Padrão:")
        defaultPromptSizer.Add(defaultPromptLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=5)
        self.defaultPromptCtrl = wx.ComboBox(generalBox.GetStaticBox(), value=config.conf["clipboardProcessor"]["selected_prompt"], choices=list(self.edited_prompts.keys()), style=wx.CB_READONLY)
        defaultPromptSizer.Add(self.defaultPromptCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        generalBox.Add(defaultPromptSizer, flag=wx.EXPAND)
        mainSizer.Add(generalBox, flag=wx.EXPAND | wx.ALL, border=5)

        # Gerenciador de Prompts
        promptsBox = wx.StaticBoxSizer(wx.VERTICAL, self, label="Gerenciador de Prompts")
        managerSizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.promptListCtrl = wx.ListBox(promptsBox.GetStaticBox(), choices=list(self.edited_prompts.keys()))
        managerSizer.Add(self.promptListCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)

        editorSizer = wx.BoxSizer(wx.VERTICAL)
        self.promptContentCtrl = wx.TextCtrl(promptsBox.GetStaticBox(), style=wx.TE_MULTILINE)
        editorSizer.Add(self.promptContentCtrl, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        
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
            self.promptContentCtrl.SetValue(self.edited_prompts.get(sel_name, ""))
        self.update_controls_state()

    def on_save_edit(self, event):
        sel_name = self.promptListCtrl.GetStringSelection()
        if not sel_name:
            return
        new_content = self.promptContentCtrl.GetValue()
        self.edited_prompts[sel_name] = new_content
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
            self.edited_prompts[name] = prompt
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

    def onSave(self):
        global prompts_collection
        config.conf["clipboardProcessor"]["api_key"] = self.apiKeyCtrl.GetValue()
        config.conf["clipboardProcessor"]["model"] = self.modelCtrl.GetValue()
        config.conf["clipboardProcessor"]["selected_prompt"] = self.defaultPromptCtrl.GetValue()
        
        save_prompts(self.edited_prompts)
        prompts_collection = copy.deepcopy(self.edited_prompts)

# --- Lógica do Plugin Global ---
class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    
    script_category = "Processador de Clipboard com IA"

    def __init__(self, *args, **kwargs):
        super(GlobalPlugin, self).__init__(*args, **kwargs)
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
        super(GlobalPlugin, self).terminate(*args, **kwargs)

    def _worker_thread(self, text_to_process, selected_prompt_name):
        api_key = config.conf["clipboardProcessor"]["api_key"]
        model = config.conf["clipboardProcessor"]["model"]
        
        system_prompt = prompts_collection.get(selected_prompt_name)
        if not system_prompt:
            wx.CallAfter(ui.message, f"Erro: Prompt '{selected_prompt_name}' não encontrado.")
            return

        if not requests:
            wx.CallAfter(ui.message, "Erro: A biblioteca 'requests' não foi encontrada.")
            return

        if not api_key:
            wx.CallAfter(ui.message, "Por favor, configure sua chave da API da OpenAI.")
            return

        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        payload = { "model": model, "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": text_to_process}] }

        try:
            response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload, timeout=30)
            response.raise_for_status()
            json_response = response.json()
            result_text = json_response["choices"][0]["message"]["content"]
            wx.CallAfter(self._update_clipboard, result_text)
        except requests.exceptions.RequestException as e:
            wx.CallAfter(ui.message, f"Erro de conexão: {e}")
        except (KeyError, IndexError):
            wx.CallAfter(ui.message, "Erro: Resposta inesperada da API.")
        except Exception as e:
            wx.CallAfter(ui.message, f"Ocorreu um erro desconhecido: {e}")

    def _update_clipboard(self, text):
        try:
            if wx.TheClipboard.Open():
                wx.TheClipboard.SetData(wx.TextDataObject(text))
                wx.TheClipboard.Close()
                tones.beep(800, 100)
                ui.message("Área de transferência atualizada pela IA.")
        except Exception as e:
            ui.message(f"Erro ao atualizar a área de transferência: {e}")

    def _start_processing(self, text_to_process, selected_prompt_name):
        ui.message(f"Processando com '{selected_prompt_name}'...")
        tones.beep(300, 100)
        thread = threading.Thread(target=self._worker_thread, args=(text_to_process, selected_prompt_name))
        thread.start()

    def _show_prompt_selection_menu(self, text_to_process):
        if not text_to_process or not text_to_process.strip():
            ui.message("Nenhum texto para processar.")
            return

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
            self._start_processing(text_to_process, selected_prompt_name)
        
        dialog.Destroy()

    def script_processSelection(self, gesture):
        try:
            focus = api.getFocusObject()
            selection = focus.selection
            selected_text = selection.text
            wx.CallAfter(self._show_prompt_selection_menu, selected_text)
        except (AttributeError, TypeError):
            ui.message("Nenhum texto selecionado.")

    def script_processClipboard(self, gesture):
        clipboard_text = ""
        try:
            if wx.TheClipboard.Open():
                if wx.TheClipboard.IsSupported(wx.DataFormat(wx.DF_TEXT)):
                    data = wx.TextDataObject()
                    if wx.TheClipboard.GetData(data):
                        clipboard_text = data.GetText()
                wx.TheClipboard.Close()
        except Exception as e:
            ui.message(f"Não foi possível ler a área de transferência: {e}")
            return
        wx.CallAfter(self._show_prompt_selection_menu, clipboard_text)

    __gestures = {
        "kb:NVDA+shift+p": "processSelection",
        "kb:control+NVDA+shift+p": "processClipboard",
    }