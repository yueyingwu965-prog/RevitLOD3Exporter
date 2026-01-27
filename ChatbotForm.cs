using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;

namespace RevitLOD3Exporter
{

    public class ChatbotForm : Form
    {

        private TextBox configSummary;

        private ListBox listAvailableTypes;
        private ListBox listSelectedTypes;
        private Button buttonTypeToRight;
        private Button buttonTypeToLeft;

        private ListBox listAvailableAttributes;
        private ListBox listSelectedAttributes;
        private Button buttonAttrToRight;
        private Button buttonAttrToLeft;

        private RichTextBox chatHistory;
        private TextBox userInput;
        private Button parseButton;
        private Button okButton;

        private readonly CityJSONData _lod2Data;

        private AiExportHelper _aiHelper;

        public ExportConfig SelectedConfig { get; private set; } = new ExportConfig();

        public ChatbotForm(CityJSONData lod2Data)
        {
            _lod2Data = lod2Data;

            // ======== 基本窗口属性 ========
            this.Text = "LOD3 Export Assistant";
            this.ClientSize = new Size(1000, 720);
            this.MinimumSize = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            // ======== 顶部：摘要 ========
            configSummary = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Top,
                Height = 140,
                ScrollBars = ScrollBars.Vertical
            };

            // ======== 中部：四列表容器 ========
            var selectionPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 320
            };

            // --- Types ---
            var lblTypeLeft = new Label
            {
                Text = "Available Types (from LOD2 CityObjects.type)",
                Left = 10,
                Top = 5,
                AutoSize = true
            };
            var lblTypeRight = new Label
            {
                Text = "Selected Types (export to LOD3)",
                Left = 500,
                Top = 5,
                AutoSize = true
            };

            listAvailableTypes = new ListBox
            {
                Left = 10,
                Top = 30,
                Width = 380,
                Height = 120,
                SelectionMode = SelectionMode.MultiExtended
            };
            listSelectedTypes = new ListBox
            {
                Left = 500,
                Top = 30,
                Width = 380,
                Height = 120,
                SelectionMode = SelectionMode.MultiExtended
            };

            buttonTypeToRight = new Button
            {
                Text = "→",
                Left = 420,
                Top = 55,
                Width = 50
            };
            buttonTypeToRight.Click += ButtonTypeToRight_Click;

            buttonTypeToLeft = new Button
            {
                Text = "←",
                Left = 420,
                Top = 95,
                Width = 50
            };
            buttonTypeToLeft.Click += ButtonTypeToLeft_Click;

            // --- Attributes ---
            var lblAttrLeft = new Label
            {
                Text = "Available Attributes (from LOD2 attributes)",
                Left = 10,
                Top = 160,
                AutoSize = true
            };
            var lblAttrRight = new Label
            {
                Text = "Selected Attributes (export to LOD3)",
                Left = 500,
                Top = 160,
                AutoSize = true
            };

            listAvailableAttributes = new ListBox
            {
                Left = 10,
                Top = 185,
                Width = 380,
                Height = 120,
                SelectionMode = SelectionMode.MultiExtended
            };
            listSelectedAttributes = new ListBox
            {
                Left = 500,
                Top = 185,
                Width = 380,
                Height = 120,
                SelectionMode = SelectionMode.MultiExtended
            };

            buttonAttrToRight = new Button
            {
                Text = "→",
                Left = 420,
                Top = 205,
                Width = 50
            };
            buttonAttrToRight.Click += ButtonAttrToRight_Click;

            buttonAttrToLeft = new Button
            {
                Text = "←",
                Left = 420,
                Top = 245,
                Width = 50
            };
            buttonAttrToLeft.Click += ButtonAttrToLeft_Click;

            // 添加到 selectionPanel
            selectionPanel.Controls.Add(lblTypeLeft);
            selectionPanel.Controls.Add(lblTypeRight);
            selectionPanel.Controls.Add(listAvailableTypes);
            selectionPanel.Controls.Add(listSelectedTypes);
            selectionPanel.Controls.Add(buttonTypeToRight);
            selectionPanel.Controls.Add(buttonTypeToLeft);

            selectionPanel.Controls.Add(lblAttrLeft);
            selectionPanel.Controls.Add(lblAttrRight);
            selectionPanel.Controls.Add(listAvailableAttributes);
            selectionPanel.Controls.Add(listSelectedAttributes);
            selectionPanel.Controls.Add(buttonAttrToRight);
            selectionPanel.Controls.Add(buttonAttrToLeft);

            // ======== 聊天记录 ========
            chatHistory = new RichTextBox
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                BackColor = Color.White,
                BorderStyle = BorderStyle.None
            };

            // ======== 聊天输入区 ========
            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40
            };

            userInput = new TextBox
            {
                Left = 5,
                Top = 8,
                Width = 650,
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
            };
            userInput.KeyDown += UserInput_KeyDown;

            parseButton = new Button
            {
                Text = "Parse",
                Width = 100,
                Height = 26,
                Anchor = AnchorStyles.Right | AnchorStyles.Top
            };
            parseButton.Click += ParseButton_Click;

            okButton = new Button
            {
                Text = "Confirm",
                Width = 100,
                Height = 26,
                Anchor = AnchorStyles.Right | AnchorStyles.Top
            };
            okButton.Click += OkButton_Click;

            bottomPanel.SizeChanged += (s, e) =>
            {
                okButton.Left = bottomPanel.ClientSize.Width - okButton.Width - 10;
                okButton.Top = 7;

                parseButton.Left = okButton.Left - parseButton.Width - 10;
                parseButton.Top = 7;

                userInput.Width = parseButton.Left - 15;
            };

            bottomPanel.Controls.Add(userInput);
            bottomPanel.Controls.Add(parseButton);
            bottomPanel.Controls.Add(okButton);

            // ======== 装载所有控件 ========
            this.Controls.Add(chatHistory);
            this.Controls.Add(bottomPanel);
            this.Controls.Add(selectionPanel);
            this.Controls.Add(configSummary);

            // ======== 在 UI 创建之后初始化 AI （重要）=======
            InitializeAIProvider();

            // ======== 聊天首条提示 ========
            AppendChat("Assistant",
                "You can either:\r\n" +
                " - use this chat (e.g., \"walls, roofs; attrs: class, revitElementId\"), or\r\n" +
                " - use the lists above to choose Types / Attributes.\r\n" +
                "Both will update the configuration at the top.");

            // ======== 初始化列表 ========
            InitializeTypeListsFromLod2();
            InitializeAttributeListsFromLod2();
            SyncConfigFromListBoxes();
            UpdateConfigSummary();
        }

        // ==================== AI Provider 初始化 ====================

        private void InitializeAIProvider()
        {
            try
            {
                string deepseek = Environment.GetEnvironmentVariable("DEEPSEEK_API_KEY");
                string gemini = Environment.GetEnvironmentVariable("GEMINI_API_KEY");
                string openai = Environment.GetEnvironmentVariable("OPENAI_API_KEY");

                if (!string.IsNullOrWhiteSpace(deepseek))
                {
                    _aiHelper = new AiExportHelper(AiProvider.DeepSeek, deepseek);
                    AppendChat("Assistant", "[AI] DeepSeek enabled.");
                }
                else if (!string.IsNullOrWhiteSpace(gemini))
                {
                    _aiHelper = new AiExportHelper(AiProvider.Gemini, gemini);
                    AppendChat("Assistant", "[AI] Gemini enabled.");
                }
                else if (!string.IsNullOrWhiteSpace(openai))
                {
                    _aiHelper = new AiExportHelper(AiProvider.OpenAI, openai);
                    AppendChat("Assistant", "[AI] OpenAI enabled.");
                }
                else
                {
                    AppendChat("Assistant",
                        "[AI] No API key found. Using rule-based parser only.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());   // 使用一下 ex，消除警告
                AppendChat("Assistant",
                    "[AI] Failed to initialize AI provider. Using rule-based parser only.");
            }
        }

        // ==================== AppendChat（已加安全保护） ====================

        private void AppendChat(string sender, string text)
        {
            if (chatHistory == null) return;

            chatHistory.SelectionStart = chatHistory.TextLength;
            chatHistory.SelectionLength = 0;

            // 判断发送者不同气泡风格
            bool isUser = sender.Equals("User", StringComparison.OrdinalIgnoreCase);

            // 气泡颜色
            Color bubbleColor = isUser ? Color.FromArgb(220, 248, 198) : Color.FromArgb(240, 240, 240);
            Color textColor = Color.Black;

            // 左右对齐
            HorizontalAlignment align = isUser ? HorizontalAlignment.Right : HorizontalAlignment.Left;

            // 插入空行（气泡间距）
            chatHistory.AppendText(Environment.NewLine);

            // 添加文本
            chatHistory.SelectionAlignment = align;
            chatHistory.SelectionBackColor = bubbleColor;
            chatHistory.SelectionColor = textColor;

            chatHistory.AppendText(text + Environment.NewLine);

            chatHistory.SelectionBackColor = chatHistory.BackColor;
            chatHistory.SelectionColor = Color.Black;

            // 光标滚动到底部
            chatHistory.SelectionStart = chatHistory.Text.Length;
            chatHistory.ScrollToCaret();
        }


        // ==================== 聊天输入 Enter 捕捉 ====================

        private void UserInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                ParseButton_Click(sender, EventArgs.Empty);
            }
        }

        // ==================== 初始化 Types ====================

        private void InitializeTypeListsFromLod2()
        {
            var allTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (_lod2Data != null && _lod2Data.CityObjects != null)
            {
                foreach (var obj in _lod2Data.CityObjects.Values)
                {
                    if (!string.IsNullOrWhiteSpace(obj.Type))
                        allTypes.Add(obj.Type);
                }
            }

            var ordered = allTypes.OrderBy(t => t).ToList();

            var selectedSet = new HashSet<string>(
                SelectedConfig.AllowedTypes ?? new HashSet<string>(),
                StringComparer.OrdinalIgnoreCase);

            var availableList = new List<string>();
            var selectedList = new List<string>();

            foreach (var t in ordered)
            {
                if (selectedSet.Contains(t))
                    selectedList.Add(t);
                else
                    availableList.Add(t);
            }

            listAvailableTypes.Items.Clear();
            listSelectedTypes.Items.Clear();

            listAvailableTypes.Items.AddRange(availableList.ToArray());
            listSelectedTypes.Items.AddRange(selectedList.ToArray());
        }

        // ==================== 初始化 Attributes ====================

        private void InitializeAttributeListsFromLod2()
        {
            var allNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (_lod2Data != null && _lod2Data.CityObjects != null)
            {
                foreach (var obj in _lod2Data.CityObjects.Values)
                {
                    if (obj.Attributes == null) continue;

                    foreach (var key in obj.Attributes.Keys)
                    {
                        allNames.Add(key);
                    }
                }
            }

            var ordered = allNames.OrderBy(n => n).ToList();

            var selectedSet = new HashSet<string>(
                SelectedConfig.SelectedAttributes ?? new List<string>(),
                StringComparer.OrdinalIgnoreCase);

            var availableList = new List<string>();
            var selectedList = new List<string>();

            foreach (var name in ordered)
            {
                if (selectedSet.Contains(name))
                    selectedList.Add(name);
                else
                    availableList.Add(name);
            }

            listAvailableAttributes.Items.Clear();
            listSelectedAttributes.Items.Clear();

            listAvailableAttributes.Items.AddRange(availableList.ToArray());
            listSelectedAttributes.Items.AddRange(selectedList.ToArray());
        }

        // ==================== 同步右侧 ListBox → Config ====================

        private void SyncConfigFromListBoxes()
        {
            var finalTypes = listSelectedTypes.Items.Cast<string>().ToList();
            SelectedConfig.AllowedTypes = new HashSet<string>(
                finalTypes, StringComparer.OrdinalIgnoreCase);

            var finalAttrs = listSelectedAttributes.Items.Cast<string>().ToList();
            SelectedConfig.SelectedAttributes = finalAttrs;
        }

        // ==================== 更新摘要 ====================

        private void UpdateConfigSummary()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Types being exported:");
            if (SelectedConfig.AllowedTypes == null || SelectedConfig.AllowedTypes.Count == 0)
            {
                sb.AppendLine(" - (All Types)");
            }
            else
            {
                foreach (var t in SelectedConfig.AllowedTypes)
                    sb.AppendLine($" - {t}");
            }

            sb.AppendLine();
            sb.AppendLine("Attributes being exported:");

            if (SelectedConfig.SelectedAttributes != null && SelectedConfig.SelectedAttributes.Count > 0)
            {
                foreach (var a in SelectedConfig.SelectedAttributes)
                    sb.AppendLine($" - {a}");
            }
            else
            {
                sb.AppendLine(" - (All Attributes)");
            }

            configSummary.Text = sb.ToString();
        }

        // ==================== Type 左右按钮 ====================

        private void ButtonTypeToRight_Click(object sender, EventArgs e)
        {
            var selected = listAvailableTypes.SelectedItems.Cast<string>().ToList();
            if (!selected.Any()) return;

            var left = listAvailableTypes.Items.Cast<string>().ToList();
            var right = listSelectedTypes.Items.Cast<string>().ToList();

            foreach (var item in selected)
            {
                if (!right.Contains(item))
                    right.Add(item);
                left.Remove(item);
            }

            listAvailableTypes.Items.Clear();
            listAvailableTypes.Items.AddRange(left.OrderBy(x => x).ToArray());

            listSelectedTypes.Items.Clear();
            listSelectedTypes.Items.AddRange(right.OrderBy(x => x).ToArray());

            SyncConfigFromListBoxes();
            UpdateConfigSummary();
        }

        private void ButtonTypeToLeft_Click(object sender, EventArgs e)
        {
            var selected = listSelectedTypes.SelectedItems.Cast<string>().ToList();
            if (!selected.Any()) return;

            var left = listAvailableTypes.Items.Cast<string>().ToList();
            var right = listSelectedTypes.Items.Cast<string>().ToList();

            foreach (var item in selected)
            {
                if (!left.Contains(item))
                    left.Add(item);
                right.Remove(item);
            }

            listAvailableTypes.Items.Clear();
            listAvailableTypes.Items.AddRange(left.OrderBy(x => x).ToArray());

            listSelectedTypes.Items.Clear();
            listSelectedTypes.Items.AddRange(right.OrderBy(x => x).ToArray());

            SyncConfigFromListBoxes();
            UpdateConfigSummary();
        }

        // ==================== Attributes 左右按钮 ====================

        private void ButtonAttrToRight_Click(object sender, EventArgs e)
        {
            var selected = listAvailableAttributes.SelectedItems.Cast<string>().ToList();
            if (!selected.Any()) return;

            var left = listAvailableAttributes.Items.Cast<string>().ToList();
            var right = listSelectedAttributes.Items.Cast<string>().ToList();

            foreach (var item in selected)
            {
                if (!right.Contains(item))
                    right.Add(item);
                left.Remove(item);
            }

            listAvailableAttributes.Items.Clear();
            listAvailableAttributes.Items.AddRange(left.OrderBy(x => x).ToArray());

            listSelectedAttributes.Items.Clear();
            listSelectedAttributes.Items.AddRange(right.OrderBy(x => x).ToArray());

            SyncConfigFromListBoxes();
            UpdateConfigSummary();
        }

        private void ButtonAttrToLeft_Click(object sender, EventArgs e)
        {
            var selected = listSelectedAttributes.SelectedItems.Cast<string>().ToList();
            if (!selected.Any()) return;

            var left = listAvailableAttributes.Items.Cast<string>().ToList();
            var right = listSelectedAttributes.Items.Cast<string>().ToList();

            foreach (var item in selected)
            {
                if (!left.Contains(item))
                    left.Add(item);
                right.Remove(item);
            }

            listAvailableAttributes.Items.Clear();
            listAvailableAttributes.Items.AddRange(left.OrderBy(x => x).ToArray());

            listSelectedAttributes.Items.Clear();
            listSelectedAttributes.Items.AddRange(right.OrderBy(x => x).ToArray());

            SyncConfigFromListBoxes();
            UpdateConfigSummary();
        }

        // ==================== Parse（AI 或 fallback） ====================

        private async void ParseButton_Click(object sender, EventArgs e)
        {
            string input = userInput.Text.Trim();
            if (string.IsNullOrEmpty(input))
                return;

            AppendChat("User", input);

            // 先同步 ListBox → Config
            SyncConfigFromListBoxes();

            bool aiUsed = false;

            // 优先尝试 AI
            if (_aiHelper != null)
            {
                try
                {
                    aiUsed = true;
                    var aiResult = await _aiHelper.ParseUserTextAsync(input);

                    // 合并结果
                    if (aiResult.Types.Count > 0)
                    {
                        if (SelectedConfig.AllowedTypes == null)
                            SelectedConfig.AllowedTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        SelectedConfig.AllowedTypes.UnionWith(aiResult.Types);
                    }
                    if (aiResult.Attributes.Count > 0)
                    {
                        if (SelectedConfig.SelectedAttributes == null)
                            SelectedConfig.SelectedAttributes = new List<string>();

                        var attrSet = new HashSet<string>(SelectedConfig.SelectedAttributes, StringComparer.OrdinalIgnoreCase);
                        attrSet.UnionWith(aiResult.Attributes);
                        SelectedConfig.SelectedAttributes = attrSet.ToList();
                    }

                    AppendChat("Assistant", "AI parsing successful. Configuration updated.");
                }
                catch (Exception ex)
                {
                    AppendChat("Assistant",
                        $"AI parsing failed, falling back to rule-based parser. ({ex.Message})");
                }
            }

            // rule-based fallback（或没启用 AI）
            if (!aiUsed)
            {
                var parsed = ExportConfigParser.ParseUserInputToConfig(input);

                if (parsed.AllowedTypes != null && parsed.AllowedTypes.Count > 0)
                {
                    if (SelectedConfig.AllowedTypes == null)
                        SelectedConfig.AllowedTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    SelectedConfig.AllowedTypes.UnionWith(parsed.AllowedTypes);
                }
                if (parsed.SelectedAttributes != null && parsed.SelectedAttributes.Count > 0)
                {
                    if (SelectedConfig.SelectedAttributes == null)
                        SelectedConfig.SelectedAttributes = new List<string>();

                    var set = new HashSet<string>(SelectedConfig.SelectedAttributes, StringComparer.OrdinalIgnoreCase);
                    set.UnionWith(parsed.SelectedAttributes);
                    SelectedConfig.SelectedAttributes = set.ToList();
                }

                AppendChat("Assistant", "Configuration updated via rule-based parser.");
            }

            // 更新 UI
            InitializeTypeListsFromLod2();
            InitializeAttributeListsFromLod2();
            SyncConfigFromListBoxes();
            UpdateConfigSummary();

            userInput.Clear();
        }

        // ==================== Confirm ====================

        private void OkButton_Click(object sender, EventArgs e)
        {
            SyncConfigFromListBoxes();
            UpdateConfigSummary();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
