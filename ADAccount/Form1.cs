using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Contexts;
using System.Security.Policy;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Animation;
using ADAccount.Properties;
using OfficeOpenXml;
using OfficeOpenXml.Style;

using Color = System.Drawing.Color;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ADAccount {
    public partial class ADAccount : Form {
		private enum WarningStateTypes {
			NULL = 0,
			NO = 1,
			YES = 2
		};

		private Settings settings;
		private List<string> listGroups;
		private List<Dictionary<string, string>> listUsers;
		private Dictionary<string, string> tempAccount;

		public ADAccount() {
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			Control.CheckForIllegalCrossThreadCalls = false;

			this.InitializeComponent();

			this.settings = new Settings();
			this.listGroups = new List<string>();
			this.listUsers = new List<Dictionary<string, string>>();

			this.LoadConfig();

			this.RadioButtonsChanged(null, null);
		}

		private string ConfigPath => $"{Application.StartupPath}\\config.json";

		private bool LoadConfig() {
			var res = false;

			if (File.Exists(this.ConfigPath)) {
				using (var file = new FileStream(this.ConfigPath, FileMode.Open)) {
					var reader = new StreamReader(file);
					var text = reader.ReadToEnd();

					if (!String.IsNullOrWhiteSpace(text)) {
						var config = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(text);

						this.chboxForbidChange.Checked = config["ForbidChange"].GetBoolean();
						this.chboxNeverExpires.Checked = config["NeverExpires"].GetBoolean();
						this.chboxEncryption.Checked = config["Encryption"].GetBoolean();
						this.chboxGenerate.Checked = config["Generate"].GetBoolean();
						this.chboxOUGroup.Checked = config["OUGroup"].GetBoolean();
						this.txboxDomain.Text = config["Domain"].GetString();
						this.txboxPasswd.Text = config["Passwd"].GetString();

						this.radioAdd.Checked = !config["Delete"].GetBoolean();
						this.radioDelete.Checked = config["Delete"].GetBoolean();

						this.settings.DeleteHome = config["DeleteHome"].GetBoolean();
						this.settings.CreateExcel = config["CreateExcel"].GetBoolean();
						this.settings.HomePath = config["HomePath"].GetString();

						res = true;
					}
				}
			}

			return res;
		}

		private void SaveConfig() {
			var config = new Dictionary<string, object>() { 
				{ "ForbidChange", this.chboxForbidChange.Checked },
				{ "NeverExpires", this.chboxNeverExpires.Checked },
				{ "Encryption", this.chboxEncryption.Checked },
				{ "Generate", this.chboxGenerate.Checked },
				{ "OUGroup", this.chboxOUGroup.Checked },
				{ "Domain", this.txboxDomain.Text },
				{ "Passwd", this.txboxPasswd.Text },
				{ "Delete", this.radioDelete.Checked },
				{ "DeleteHome", this.settings.DeleteHome },
				{ "CreateExcel", this.settings.CreateExcel },
				{ "HomePath", this.settings.HomePath }
			};

			using (var file = new FileStream(this.ConfigPath, FileMode.Create, FileAccess.ReadWrite)) {
				var writer = new StreamWriter(file, Encoding.UTF8);
				var json = JsonSerializer.Serialize(config, new JsonSerializerOptions() { WriteIndented = true });

				writer.BaseStream.Seek(0, SeekOrigin.End);
				writer.Write(json);
				writer.Flush();
				writer.Close();
				file.Close();
			}
		}

        private async Task RefreshTable() {
			var i = -1;

			this.table.Rows.Clear();

            foreach (var user in this.listUsers) {
				if (!string.IsNullOrWhiteSpace(user["logon"])) {
					this.table.Rows.Add();
					i++;
				}
			}

            var stack = new Stack<Dictionary<string, string>>(this.listUsers);

			while (stack.Count() > 0) {
				var user = stack.Pop();

				// if (!string.IsNullOrWhiteSpace(user["logon"])) {
					this.table.Rows[i].Cells[0].Value = user["fullname"];
					this.table.Rows[i].Cells[1].Value = user["group"];
					this.table.Rows[i].Cells[2].Value = $"{Regex.Split(user["dob"], " ").First()}";
					this.table.Rows[i].Cells[3].Value = user["logon"];
					i--;
				// }
			}
		}

		private async Task ClearTable() {
			this.listGroups.Clear();
			this.listGroups.Add("Без группы");
			this.listUsers.Clear();
			this.table.Rows.Clear();
		}

		private bool CheckUserList() {
			if (this.listUsers.Count() == 0) {
				MessageBox.Show("Пожалуйста, загрузите таблицу Excel.");
				return false;
			} else 
				return true;
		}

		private bool CheckDomain(string dom) {
			var exist = false;

			// Проверка на заполненное поле
			if (string.IsNullOrWhiteSpace(dom)) {
				MessageBox.Show("Пожалуйста, укажите домен.");
				return false;
			}

			// Проверка на существование доменов
			foreach (Domain item in Forest.GetCurrentForest().Domains) {
				if (string.Compare(item.Name.ToLower(), dom.ToLower()) == 0) {
					exist = true;
					break;
				}
			}

			if (!exist) {
				MessageBox.Show("Домен был неверно указан или не существует.");
				return false;
			}

			return true;
		}

		private bool CheckOU(string path) {
			try {
				var entry = new DirectoryEntry(path);
				var guid = entry.Guid;
				return true;
			} catch {
				return false;
			}
		}

		private string GetLDAP(string src) {
			var path = Regex.Split(src, "/");
			var args = string.Empty;
			var dom = path.First();

			if (path.Count() > 1) {
				for (int i = path.Count() - 1; i > 0; i--)
					args += $"OU={path[i]},";
			}

			var subs = dom.Split('.');

			foreach (var sub in subs)
				args += $"DC={sub}" + (sub != subs.Last() ? "," : "");

			return $"LDAP://127.0.0.1/{args}";
		}

		private void CreateTemporaryAdminAccount(PrincipalContext context) {
			var dialog = DialogResult.Retry;

			while (dialog == DialogResult.Retry) {
				try {
					var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
					var rand = new Random();
					var name = "ADA-";
					var passwd = string.Empty;

					for (int i = 0; i < 12; i++)
						name += chars[rand.Next(chars.Length)];

					for (int i = 0; i < 16; i++)
						passwd += chars[rand.Next(chars.Length)];

					var up = new UserPrincipal(context);
					up.Name = up.DisplayName = up.SamAccountName = up.UserPrincipalName = name;
					up.Enabled = true;
					up.SetPassword(passwd);
					up.Save();

					this.tempAccount = new Dictionary<string, string>() { { "logon", name }, 
																		  { "passwd", passwd } };

					var group = GroupPrincipal.FindByIdentity(context, "Администраторы домена");

					group.Members.Add(up);
					group.Save();
					break;
				} catch (Exception ex) {
					dialog = MessageBox.Show(ex.ToString(), this.Text, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
				}
			}
		}

		private void DeleteTemporaryAdminAccount(PrincipalContext context) {
			var dialog = DialogResult.Retry;

			while (dialog == DialogResult.Retry) {
				try {
					var up = UserPrincipal.FindByIdentity(context, this.tempAccount["logon"]);
					up.Delete();

					this.tempAccount["logon"] = string.Empty;
					this.tempAccount["passwd"] = string.Empty;
					break;
				} catch (Exception ex) {
					dialog = MessageBox.Show(ex.ToString(), this.Text, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
				}
			}
		}

		private PrincipalContext InitPrincipalContext(string src) {
			PrincipalContext context = null;

			var path = Regex.Split(src, "/");
			var dom = path.First();

			if (this.CheckDomain(dom)) {
				var args = string.Empty;

				if (path.Count() > 1) {
					for (int i = path.Count() - 1; i > 0; i--)
						args += $"OU={path[i]},";

					var subs = dom.Split('.');

					foreach (var sub in subs)
						args += $"DC={sub}" + (sub != subs.Last() ? "," : "");

					if (!this.CheckOU($"LDAP://127.0.0.1/{args}")) {
						MessageBox.Show($"LDAP://127.0.0.1/{args}\n\nУказанные OU не существуют в домене \"{dom}\".");
					} else {
						context = new PrincipalContext(ContextType.Domain, dom, args);
					}
				} else {
					context = new PrincipalContext(ContextType.Domain, dom);
				}
			} 
			return context;
		}

        private void AddDomainUsers() {
			DialogResult dialog;

			try {
				if (!this.CheckUserList())
					return;

				var ignore = new List<string>();
				var list = new List<string>();
				var addr = this.txboxDomain.Text;

				// Создаем пользователя домена
				foreach (var user in this.listUsers) {
					var context = this.InitPrincipalContext(addr);

					if (context == null)
						return;

					// Добавляем пользователя в OU, если группа была указана в таблице
					if (this.chboxOUGroup.Checked) {
						foreach (var item in Regex.Split(user["group"], ",")) {
							var group = item.Trim();
							var path = this.GetLDAP($"{addr}/{group}");
							
							if (!this.CheckOU(path)) {
								var skip = false;

								foreach (var ig in ignore) {
									if (string.Compare(ig, group) == 0)
										skip = true;
								}
								if (skip)
									continue;

								if (!string.IsNullOrWhiteSpace(group)) { 
									dialog = MessageBox.Show($"OU подразделение \"{group}\" не существует. Создать новое OU подразделение?",
															 this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
									if (dialog == DialogResult.Yes) {
										var root = this.InitPrincipalContext(addr.Split('/').First());

										this.CreateTemporaryAdminAccount(root);

										var obj_adam = new DirectoryEntry(this.GetLDAP($"{addr}"));
                                        obj_adam.Username = this.tempAccount["logon"];
                                        obj_adam.Password = this.tempAccount["passwd"];
                                        obj_adam.RefreshCache();

										var obj_ou = obj_adam.Children.Add($"OU={group}", "OrganizationalUnit");
                                        obj_ou.Properties["description"].Add(group);
                                        obj_ou.CommitChanges();

										this.DeleteTemporaryAdminAccount(root);
									} else {
										ignore.Add(group);
									}
								}
							}
							if (this.CheckOU(path))
								context = this.InitPrincipalContext($"{addr}/{group}");
						}
					}
					var up = new UserPrincipal(context);
					up.Name = user["fullname"];
					up.GivenName = user["name"];
					up.Surname = user["surname"];
                    up.DisplayName = user["fullname"];
					up.UserPrincipalName = user["logon"];
					up.SamAccountName = user["logon"];
					up.UserCannotChangePassword = this.chboxForbidChange.Checked;
					up.PasswordNeverExpires = this.chboxNeverExpires.Checked;
                    up.AllowReversiblePasswordEncryption = this.chboxEncryption.Checked;
					up.PasswordNotRequired = false;
					up.Enabled = true;

					if (this.chboxGenerate.Checked) {
						DateTime dob;

						if (!DateTime.TryParse(user["dob"], out dob)) {
							MessageBox.Show($"Пользователь \"{user["fullname"]}\" не имеет дату рождения. " +
											 "Невозможно сгенерировать пароль по установленным правилам.", this.Text,
											 MessageBoxButtons.OK, MessageBoxIcon.Error);
							continue;
						} else {
							var surname = user["surname"];
							var passwd = surname.Count() < 3 ? surname.ToLower() : surname.Substring(0, 3).ToLower();

							while (passwd.Count() < 4) {
                                passwd += '_';
                            }

							up.SetPassword(user["passwd"] = passwd += $"{dob.Day:D2}{dob.Month:D2}");
						}
					} else if (!string.IsNullOrWhiteSpace(this.txboxPasswd.Text)) {
						up.SetPassword(user["passwd"] = this.txboxPasswd.Text);
					} else {
						up.PasswordNotRequired = true;
					}
					up.Save();
					list.Add(user["fullname"]);

					// Добавляем пользователя в группу, если была указана в таблице
					if (!string.IsNullOrWhiteSpace(user["group"])) {
						foreach (var item in Regex.Split(user["group"], ",")) {
							if (!this.chboxOUGroup.Checked) {
								var group = GroupPrincipal.FindByIdentity(context, item.Trim());

								if (group == null) {
									var skip = false;

									// Чтобы не задавались одни и те же вопросы, пропускаем
									foreach (var ig in ignore) {
										if (string.Compare(ig, user["group"]) == 0)
											skip = true;
									}
									dialog = MessageBox.Show($"Группа пользователей \"{user["group"]}\" не существует. " +
															 $"Создать новую группу?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

									if (dialog == DialogResult.Yes) {
										group = new GroupPrincipal(context, user["group"]);
									} else {
										ignore.Add(user["group"]);
										continue;
									}
								}
								group.Members.Add(up);
								group.Save();
							}
						}
					}
				}
 
				if (list.Count() != 0) {
					if (this.settings.CreateExcel) {
						dialog = MessageBox.Show("Создать таблицу Excel пользователей (логин, пароль)?", 
												 this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
						if (dialog == DialogResult.Yes) {
							var save = new SaveFileDialog() { Filter = "Файл Excel (*.xlsx, *.xls)|*.xlsx;*.xls" };
							dialog = save.ShowDialog();

							if (dialog == DialogResult.OK) {
								this.SaveExcel(save.FileName);
							}
						}
					} 
					MessageBox.Show($"Вами были добавлены пользователи домена: \n\n{string.Join("\n", list)}", 
									this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
				} else {
					MessageBox.Show("Не были найдены пользователи для добавления.", this.Text, 
									MessageBoxButtons.OK, MessageBoxIcon.Information);
				}

				this.ClearTable();
			} catch (Exception ex) {
				MessageBox.Show(ex.ToString(), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
        }

		private void DeleteDomainUsers() {
			DialogResult		dialog;
			WarningStateTypes   warning = WarningStateTypes.NULL;

            if (this.listUsers.Count() == 0) {
				MessageBox.Show("Пожалуйста, загрузите таблицу Excel.");
				return;
			}

			if (string.IsNullOrEmpty(this.settings.HomePath)) {
                MessageBox.Show(
					"Не могу удалить домашние директории пользователей. Укажите путь домашних директорий в настройках.",
					this.Text,
					MessageBoxButtons.OK, 
					MessageBoxIcon.Warning
				);
				return;
            }

			try {
				if (!this.CheckUserList())
					return;

				var ignore = new List<string>() { "Пользователи домена" };
				var list = new List<string>();
				var addr = this.txboxDomain.Text;
				var context = this.InitPrincipalContext(addr);

                if (context == null)
                    return;

                foreach (var user in this.listUsers) {
					var up = UserPrincipal.FindByIdentity(context, user["fullname"]);

					if (up == null) {
						if (warning != WarningStateTypes.NO) {
							MessageBox.Show($"Пользователь \"{user["fullname"]}\" отсутствуют в домене \"{this.txboxDomain.Text}\".", this.Text,
											MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
                        if (warning == WarningStateTypes.NULL) {
                            dialog = MessageBox.Show($"Предупреждать, если пользователи отсутствуют в домене?",
                                                     this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            warning = dialog == DialogResult.Yes ? WarningStateTypes.YES : WarningStateTypes.NO;
                        }
                        break;

                        MessageBox.Show($"Пользователь \"{user["fullname"]}\" не был найден в домене.");
						continue;
					}

					var logon = up.SamAccountName;
					var back = false;

					list.Add(user["fullname"]);

					if (!chboxOUGroup.Checked) {
                        // Удаление группы
                        foreach (GroupPrincipal group in up.GetGroups()) {
							if (string.IsNullOrWhiteSpace(group.SamAccountName))
								continue;

							foreach (var sam in ignore) {
								if (group.SamAccountName == sam)
									goto next;
							}

							if (group.Members.Count() <= 1) {
								dialog = MessageBox.Show($"В группе \"{group.SamAccountName}\" не осталось пользователей. " +
														 $"Удалить пустую группу?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

								if (dialog == DialogResult.Yes)
									group.Delete();
							}
						next:
							continue;
						}
					} else {
                        // FIXME: Должно быть удаление OU, но не могу подобрать хорошую реализацию
                    }
                    up.Delete();

					// Удаление домашней директории
					if (this.settings.DeleteHome) {
						var paths = Directory.GetDirectories(this.settings.HomePath);

						foreach (var path in paths) {
							var dir = Regex.Split(path, "\\\\").Last();

							if (string.Compare(logon, dir) == 0) {
								if (Directory.Exists(path)) {
									while (true) {
										try {
											Directory.Delete(path, true);
											break;
										} catch (Exception ex) {
											dialog = MessageBox.Show(ex.ToString(), this.Text, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);

											if (dialog != DialogResult.Retry)
												break;
										} break;
									}
								}
							}
						}
					}
				} 
				
				MessageBox.Show(list.Count() != 0 ? $"Вами были удалены пользователи домена: \n\n{string.Join("\n", list)}" :
								"Не были найдены пользователи для удаления.", 
								this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.ClearTable();
			} catch (Exception ex) {
				MessageBox.Show(ex.ToString(), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void LoadExcel(string path) {
            var warning = WarningStateTypes.NULL;
            var context = this.InitPrincipalContext(this.txboxDomain.Text);

			if (context == null)
				return;

			try {
				if (string.Compare(path, path.Length - 5, ".xls", 0, 4) == 0);
				else if (string.Compare(path, path.Length - 6, ".xlsx", 0, 5) == 0);
				else return;
			} catch {
				return; 
			}
			var package = new ExcelPackage(new FileInfo(path));
            var worksheet = package.Workbook.Worksheets[0];
            var cells = worksheet.Cells;

			DateTime date;

			this.ClearTable().Wait();

			for (int i = 1; cells[i, 1].Value != null; i++) {
				if (string.Compare("Фамилия", $"{cells[i, 1].Value}") != 0);
				else if (string.Compare("ФИО", $"{cells[i, 1].Value}") != 0);
				else continue;

				var user1 = new Dictionary<string, string>() {
                    { "name", string.Empty },
					{ "surname", string.Empty },
					{ "fullname", string.Empty },
					{ "dob", string.Empty },
					{ "group", string.Empty },
                    { "logon", string.Empty },
					{ "passwd", string.Empty }
				};

				for (int k = 1; k <= 5; k++) {
					if (k == 1) {
						user1["surname"] = $"{cells[i, k].Value}";
					} else if (k == 2) {
						user1["name"] = $"{cells[i, k].Value}";
					}
					if (k < 4) {
						user1["fullname"] += $"{cells[i, k].Value}" + (k >= 3 ? "" : " ");
					} else if (k < 5) {
						// На основе полученных данных о ФИО, генерируем и проверяем логин
						UserPrincipal up;
						var name1 = Regex.Split(user1["fullname"], " ");

						if (this.radioAdd.Checked) {
							var add_prefix = false;

							up = UserPrincipal.FindByIdentity(context, user1["surname"]);

							if (up != null)
								add_prefix = true;

							if (name1.Length >= 3) {
								if (UserPrincipal.FindByIdentity(context, user1["fullname"]) != null) {
									if (warning != WarningStateTypes.NO)
										MessageBox.Show($"Пользователь \"{user1["fullname"]}\" был создан ранее.", this.Text,
														MessageBoxButtons.OK, MessageBoxIcon.Warning);

									if (warning == WarningStateTypes.NULL) {
										var dialog = MessageBox.Show($"Предупреждать, если пользователи уже существуют в домене?",
																	 this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

										warning = dialog == DialogResult.Yes ? WarningStateTypes.YES : WarningStateTypes.NO;
									}
									break;
								}
								foreach (var user2 in this.listUsers) {
									if (add_prefix)
										break;

									var name2 = Regex.Split(user2["fullname"], " ");

									if (name2.Length >= 3) {
										if (name1[0] == name2[0])
											add_prefix = true;
									}
								}
								user1["logon"] = add_prefix ? $"{name1[0]}_{name1[1][0]}{name1[2][0]}" : name1[0];
							} else if (name1.Length == 1)
								user1["logon"] = name1[0];
						} else if (this.radioDelete.Checked) {
							up = UserPrincipal.FindByIdentity(context, user1["fullname"]);

							if (up == null) {
								if (warning != WarningStateTypes.NO) {
									MessageBox.Show($"Пользователь \"{user1["fullname"]}\" отсутствует в домене \"{this.txboxDomain.Text}\".", this.Text,
													MessageBoxButtons.OK, MessageBoxIcon.Warning);
								}
								if (warning == WarningStateTypes.NULL) {
									var dialog = MessageBox.Show($"Предупреждать, если пользователи не существуют в домене?",
																 this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

									warning = dialog == DialogResult.Yes ? WarningStateTypes.YES : WarningStateTypes.NO;
								}
								break;
							} else {
								user1["logon"] = up.SamAccountName;
							}
						}

						if (DateTime.TryParse($"{cells[i, k].Value}", out date)) {
							user1["dob"] = $"{date}";
						} else {
							MessageBox.Show($"У пользователя \"{user1["fullname"]}\" была неправильно указана дата рождения.",
											this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            this.listUsers.Clear();
                            return;
						}
					} else {
						var add_group = true;
						user1["group"] = $"{cells[i, k].Value}";

						if (!string.IsNullOrWhiteSpace(user1["group"])) {
							foreach (var group in this.listGroups) {
								if (group == user1["group"]) {
									add_group = false;
									break;
								}
							}
							if (add_group)
								this.listGroups.Add(user1["group"]);
						}
						this.listUsers.Add(user1);
					}
					// Проверка на повторяющееся ФИО
					if (this.radioAdd.Checked) {
						foreach (var firstUser in this.listUsers) {
							var firstFullname = firstUser["fullname"];

							foreach (var secondUser in this.listUsers) {
								var secondFullname = secondUser["fullname"];

								if (!string.ReferenceEquals(firstFullname, secondFullname)) {
									if (string.Compare(firstFullname, secondFullname) == 0) {
										MessageBox.Show($"ФИО \"{firstUser["fullname"]}\" не уникально. Проверьте таблицу Excel.",
														this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
										this.listUsers.Clear();
										return;
									}
								}
							}
						}
					}
                }
            }
			this.RefreshTable();
		}

		private void SaveExcel(string path) {
			Action<Border> border = (br) => {
				br.Left.Style = ExcelBorderStyle.Thin;
				br.Right.Style = ExcelBorderStyle.Thin;
				br.Top.Style = ExcelBorderStyle.Thin;
				br.Bottom.Style = ExcelBorderStyle.Thin;
			};

			var package = new ExcelPackage();

			foreach (var group in this.listGroups) {
				var worksheet = package.Workbook.Worksheets.Add(group);
				var cols = worksheet.Columns;
				var cells = worksheet.Cells;

				cols[1].Width = 35;
				cols[2].Width = 15;
				cols[3].Width = 15;

				cells[1, 1].Value = group;
				cells[1, 1].Style.Font.Bold = true;
				cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				cells["A1:C1"].Merge = true;
				border(cells["A1:C1"].Style.Border);

				cells[2, 1].Value = "ФИО";
				cells[2, 2].Value = "Логин";
				cells[2, 3].Value = "Пароль";
				cells["A2:C2"].Style.Font.Bold = true;
				cells["A2:C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				border(cells["A2:C2"].Style.Border);
			}

			var queue = new Queue<Dictionary<string, string>>(this.listUsers);
			var counts = new Dictionary<string, int>();

			foreach (var group in this.listGroups)
				counts.Add(group, 3);

			while (queue.Count() > 0) {
				var dict = queue.Dequeue();

				if (string.IsNullOrWhiteSpace(dict["logon"]))
					continue;

				var group = !string.IsNullOrWhiteSpace(dict["group"]) ? dict["group"] : "Без группы";
				var worksheet = package.Workbook.Worksheets[group];
				var cells = worksheet.Cells;

				cells[counts[group], 1].Value = dict["fullname"];
				border(cells[counts[group], 1].Style.Border);
				cells[counts[group], 2].Value = dict["logon"];
				border(cells[counts[group], 2].Style.Border);
				cells[counts[group], 3].Value = dict["passwd"];
				border(cells[counts[group], 3].Style.Border);
				counts[group]++;
			}

			while (true) {
				try {
					package.SaveAs(path);
					break;
				} catch {
					var dialog = MessageBox.Show($"Файл \"{path}\" не может быть сохранен.", this.Text,
												 MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);

					if (dialog != DialogResult.Retry)
						break;
				}
			}

			this.SaveConfig();
		}

		private void OpenExcelFile(object sender, EventArgs e) {
			var open = new OpenFileDialog() { Filter = "Файл Excel (*.xlsx, *.xls)|*.xlsx;*.xls" };
			var result = open.ShowDialog();

			if (result == DialogResult.OK)
				this.LoadExcel(open.FileName);
		}

		private void ResizeWindow(object sender, EventArgs e) {

		}

		private void DragEnterWindow(object sender, DragEventArgs e) {
			if (e.Data.GetDataPresent(DataFormats.FileDrop)) 
				e.Effect = DragDropEffects.Copy;
		}

		private void DragDropWindow(object sender, DragEventArgs e) {
			this.LoadExcel(((string[])e.Data.GetData(DataFormats.FileDrop)).Last());
		}

		private void Settings(object sender, EventArgs e) {
			this.settings.ShowDialog();
		}

		private void checkBox1_CheckedChanged(object sender, EventArgs e) {
			this.txboxPasswd.Enabled = !this.chboxGenerate.Checked;
		}

		private void RadioButtonsChanged(object sender, EventArgs e) {
			this.ClearTable();

			if (this.radioAdd.Checked) {
				this.chboxForbidChange.Enabled = true;
				this.chboxNeverExpires.Enabled = true;
				this.chboxEncryption.Enabled = true;
				this.chboxGenerate.Enabled = true;
				this.txboxPasswd.Enabled = !this.chboxGenerate.Checked;

				if (!this.table.Columns[1].Visible) {
					this.table.Columns[0].Width -= this.table.Columns[1].Width;
					this.table.Columns[1].Visible = true;
				}
			} else if (this.radioDelete.Checked) {
				this.chboxForbidChange.Enabled = false;
				this.chboxNeverExpires.Enabled = false;
				this.chboxEncryption.Enabled = false;
				this.chboxGenerate.Enabled = false;
				this.txboxPasswd.Enabled = false;

				if (this.table.Columns[1].Visible) {
					this.table.Columns[0].Width += this.table.Columns[1].Width;
					this.table.Columns[1].Visible = false;
				}
			}
		}

		private void ApplyChanges(object sender, EventArgs e) {
			if (this.radioAdd.Checked)
				this.AddDomainUsers();
			else if (this.radioDelete.Checked)
				this.DeleteDomainUsers();
		}

		private void pictureBox1_MouseHover(object sender, EventArgs e) {
			this.pictureBox1.Image = Properties.Resources.settings_on;
		}

		private void pictureBox1_MouseLeave(object sender, EventArgs e) {
			this.pictureBox1.Image = Properties.Resources.settings;
		}
	}

	public sealed class Settings : Form {
		private CheckBox chboxDeleteHome;
		private CheckBox chboxCreateExcel;
		private Label lblHomePath;
		private TextBox txboxHomePath;
		private Button btnReviewHome;

		public bool DeleteHome { get => this.chboxDeleteHome.Checked;
								 set => this.chboxDeleteHome.Checked = value; }
		public bool CreateExcel { get => this.chboxCreateExcel.Checked;
								  set => this.chboxCreateExcel.Checked = value; }
		public string HomePath { get => this.txboxHomePath.Text;
								 set => this.txboxHomePath.Text = value; }

		public Settings() {
            var resources = new ComponentResourceManager(typeof(ADAccount));

            this.Icon = ((Icon)(resources.GetObject("$this.Icon")));
            this.Text = "Настройки";
			this.Size = new Size(680, 240);
			this.MinimizeBox = false;
			this.MaximizeBox = false;
			this.StartPosition = FormStartPosition.CenterScreen;
			this.FormBorderStyle = FormBorderStyle.Fixed3D;

			this.chboxDeleteHome = new CheckBox() {
				Checked = true,
				Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204))),
				Text = "Удалять домашнюю директорию пользователей",
				Size = new Size(600, 32),
				Location = new Point(20, 20)
			};

			this.chboxCreateExcel = new CheckBox() {
				Checked = true,
				Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204))),
				Text = "Создавать таблицу Excel пользователей (логин, пароль)",
				Size = new Size(600, 32),
				Location = new Point(20, 52)
			};

			this.lblHomePath = new Label() {
				Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204))),
				Text = "Путь до домашней директории: ",
				Size = new Size(260, 48),
				Location = new Point(20, 92)
			};

			this.txboxHomePath = new TextBox() {
				Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204))),
				Size = new Size(240, 48),
				Location = new Point(this.lblHomePath.Size.Width + 20, this.lblHomePath.Location.Y)
			};

			this.btnReviewHome = new Button() {
				Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204))),
				Text = "Обзор",
				Size = new Size(96, 28),
				Location = new Point(this.lblHomePath.Size.Width + this.txboxHomePath.Size.Width + 30, this.lblHomePath.Location.Y - 1)
			};

			this.chboxDeleteHome.CheckedChanged += new EventHandler(this.CheckDeleteHome);
			this.btnReviewHome.Click += new EventHandler(this.ReviewHome);

			this.Controls.Add(this.chboxDeleteHome);
			this.Controls.Add(this.chboxCreateExcel);
			this.Controls.Add(this.lblHomePath);
			this.Controls.Add(this.txboxHomePath);
			this.Controls.Add(this.btnReviewHome);
		}

		private void ReviewHome(object sender1, EventArgs e1) {
			var open = new FolderBrowserDialog();
			var result = open.ShowDialog();

			if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(open.SelectedPath)) {
				this.txboxHomePath.Text = open.SelectedPath;
			}
		}

		private void CheckDeleteHome(object sender1, EventArgs e1) {
			this.txboxHomePath.ReadOnly = !this.chboxDeleteHome.Checked;
		}
	}
}
