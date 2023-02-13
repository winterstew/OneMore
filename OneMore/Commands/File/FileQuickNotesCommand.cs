﻿//************************************************************************************************
// Copyright © 2023 Steven M Cohn.  All rights reserved.
//************************************************************************************************

namespace River.OneMoreAddIn.Commands
{
	using River.OneMoreAddIn.Models;
	using River.OneMoreAddIn.Settings;
	using System;
	using System.Globalization;
	using System.Linq;
	using System.Text.RegularExpressions;
	using System.Threading.Tasks;
	using System.Xml.Linq;


	internal class FileQuickNotesCommand : Command
	{
		private OneNote one;
		private bool titled;
		private bool stamped;


		public FileQuickNotesCommand()
		{
		}


		public override async Task Execute(params object[] args)
		{
			var collection = new SettingsProvider().GetCollection(nameof(QuickNotesSheet));
			titled = collection.Get("titled", false);
			stamped = collection.Get("stamped", false);

			var organization = collection.Get("organization", "notebook");
			if (organization == "notebook")
			{
				var notebookID = collection.Get<string>("notebookID");
				if (!string.IsNullOrWhiteSpace(notebookID))
				{
					var grouping = collection.Get("grouping", 0);
					await FileIntoNotebook(notebookID, grouping);
				}
			}
			else
			{
				var sectionID = collection.Get<string>("sectionID");
				if (!string.IsNullOrWhiteSpace(sectionID))
				{
					await FileIntoSection(sectionID);
				}
			}
		}


		private async Task FileIntoNotebook(string notebookID, int grouping)
		{
			one = new OneNote();

			var unfiled = await LoadQuickNotes();
			if (unfiled == null)
			{
				return;
			}

			var notebook = await one.GetNotebook(notebookID, OneNote.Scope.Sections);
			var ns = notebook.GetNamespaceOfPrefix(OneNote.Prefix);
			var count = 0;
			string sectionID = null; // keep track of last section used

			unfiled.Descendants(ns + "Page").ForEach(async e =>
			{
				e.GetAttributeValue("name", out var name, string.Empty);
				e.GetAttributeValue("dateTime", out var dateTime);
				e.GetAttributeValue("lastModifiedTime", out var lastModifiedTime);

				var page = one.GetPage(e.Attribute("ID").Value, OneNote.PageDetail.All);
				var section = await FindFilingSection(notebook, grouping, page, DateTime.Parse(dateTime));
				sectionID = section.Attribute("ID").Value;

				AddHeader(page, name, dateTime);

				logger.WriteLine($"moving quick note [{name}] to section [{section.Attribute("name").Value}]");
				var pageID = await CopyPage(page, sectionID);
				Timewarp(sectionID, pageID, dateTime, lastModifiedTime);
				count++;
			});

			if (count > 0 && sectionID != null)
			{
				EmptyQuickNotes(unfiled);
				await one.NavigateTo(sectionID);
			}
		}


		private async Task<XElement> FindFilingSection(
			XElement notebook, int grouping, Page page, DateTime dateTime)
		{
			string name = null;
			switch (grouping)
			{
				// work week (2023-02-13 Wk 7)
				case 0:
					var week = new GregorianCalendar(GregorianCalendarTypes.Localized)
						.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

					name = $"{dateTime:yyyy-MM-dd} W{week}";
					break;

				// 1:month (2023-02)
				case 1:
					name = dateTime.ToString("yyyy-MM");
					break;

				// quarter (2023 Q1)
				case 2:
					name = $"{dateTime:yyyy} Q{(dateTime.Month + 2) / 3}";
					break;

				// year (2023)
				case 3:
					name = dateTime.ToString("yyyy");
					break;

				// #keyword
				case 4:
					var text = page.Root.Descendants(page.Namespace + "OE").First()?.TextValue();
					if (!string.IsNullOrWhiteSpace(text))
					{
						var match = Regex.Match(text, @"#[^\s]+");
						if (match.Success)
						{
							name = match.Value.Substring(1).Replace("/", "|");
							break;
						}
					}
					name = dateTime.ToString("yyyy-MM");
					break;
			}

			var ns = notebook.GetNamespaceOfPrefix(OneNote.Prefix);
			var section = notebook.Elements(ns + "Section")
				.FirstOrDefault(e => e.Attribute("name").Value == name);
			
			if (section == null)
			{
				section = new XElement(ns + "Section", new XAttribute("name", name));

				var bin = notebook.Elements(ns + "SectionGroup")
					.FirstOrDefault(e => e.Attribute("isRecycleBin") != null);

				if (bin == null)
					notebook.Add(section);
				else
					bin.AddBeforeSelf(section);

				one.UpdateHierarchy(notebook);
				var book = await one.GetNotebook(notebook.Attribute("ID").Value);
				section = book.Elements(ns + "Section")
					.First(e => e.Attribute("name").Value == name);

				notebook.Elements(ns + "Section")
					.First(e => e.Attribute("name").Value == name)
					.ReplaceWith(section);
			}

			return section;
		}


		private async Task FileIntoSection(string sectionID)
		{
			one = new OneNote();

			var unfiled = await LoadQuickNotes();
			if (unfiled == null)
			{
				return;
			}

			var section = one.GetSection(sectionID);
			var ns = section.GetNamespaceOfPrefix(OneNote.Prefix);
			var count = 0;

			unfiled.Descendants(ns + "Page").ForEach(async e =>
			{
				e.GetAttributeValue("name", out var name, string.Empty);
				e.GetAttributeValue("dateTime", out var dateTime);
				e.GetAttributeValue("lastModifiedTime", out var lastModifiedTime);

				logger.WriteLine($"moving quick note [{name}]");

				var page = one.GetPage(e.Attribute("ID").Value, OneNote.PageDetail.All);

				AddHeader(page, name, dateTime);
				var pageID = await CopyPage(page, sectionID);

				Timewarp(sectionID, pageID, dateTime, lastModifiedTime);
				count++;
			});

			if (count > 0)
			{
				EmptyQuickNotes(unfiled);
				await one.NavigateTo(sectionID);
			}
		}


		private async Task<XElement> LoadQuickNotes()
		{
			var books = await one.GetNotebooks();
			var ns = books.GetNamespaceOfPrefix(OneNote.Prefix);

			// Quick Notes are stored in the singular one:UnfiledNotes notebook node
			var book = books.Elements(ns + "UnfiledNotes").FirstOrDefault();
			if (book == null)
			{
				return null;
			}

			// get the notebook with pages...
			var unfiled = await one.GetNotebook(book.Attribute("ID").Value, OneNote.Scope.Pages);
			if (unfiled == null || !unfiled.Elements().Any())
			{
				return null;
			}

			return unfiled;
		}


		private void AddHeader(Page page, string name, string dateTime)
		{
			if (titled)
			{
				// extract text from body, possibly removing #keyword first line
				if (Regex.IsMatch(name, @"^#[^\s]+$"))
				{
					// skip first line and grab name from remaining content
					var next = page.Root.Descendants(page.Namespace + "OE")
						.Skip(1)
						.FirstOrDefault(e => e.TextValue().Length != 0);

					if (next != null)
					{
						name = next.Value;
						if (name.Length > 20)
						{
							name = name.Substring(20) + "...";
						}
					}
				}
			}
			else
			{
				// do not extract text from body
				name = string.Empty;
			}

			if (stamped && DateTime.TryParse(dateTime, out var dttm))
			{
				name = $"{dttm:yyyy-MM-dd} {name}";
			}

			page.SetTitle(name);

			// shift content down...

			page.Root.Elements(page.Namespace + "Outline").ForEach(e =>
			{
				var position = e.Element(page.Namespace + "Position");
				if (position != null)
				{
					position.GetAttributeValue("y", out var y, 0.0);
					if (y < page.TopOutlinePosition)
					{
						position.SetAttributeValue("y", $"{page.TopOutlinePosition}.0");
					}
				}
			});
		}


		private async Task<string> CopyPage(Page page, string sectionID)
		{
			one.CreatePage(sectionID, out var pageID);

			// set the page ID to the new page's ID
			page.Root.Attribute("ID").Value = pageID;
			// remove all objectID values and let OneNote generate new IDs
			page.Root.Descendants().Attributes("objectID").Remove();
			var copy = new Page(page.Root); // reparse to refresh PageId

			await one.Update(copy);

			return pageID;
		}


		private void Timewarp(
			string sectionID, string pageID, string dateTime, string lastModifiedTime)
		{
			var section = one.GetSection(sectionID);
			var ns = section.GetNamespaceOfPrefix(OneNote.Prefix);

			var page = section.Descendants(ns + "Page")
				.FirstOrDefault(e => e.Attribute("ID")?.Value == pageID);

			if (page != null)
			{
				page.SetAttributeValue("dateTime", dateTime);
				page.SetAttributeValue("lastModifiedTime", lastModifiedTime);

				one.UpdateHierarchy(section);
			}
		}


		private void EmptyQuickNotes(XElement unfiled)
		{
			var ns = unfiled.GetNamespaceOfPrefix(OneNote.Prefix);

			// need to delete individual pages in Quick Notes section because that
			// section is considered read-only by OneNote so would throw and exception

			unfiled.Descendants(ns + "Page").ForEach(e =>
			{
				one.DeleteHierarchy(e.Attribute("ID").Value);
			});
		}
	}
}
/*
<one:UnfiledNotes xmlns:one="" ID="{}">
	<one:Section name="Quick Notes" ID="{}" path="C:\Users\..\OneDrive\Documents\OneNote Notebooks\Quick Notes.one" lastModifiedTime="" color="#B7C997">
	<one:Page ID="{}" name="This is a quick note" dateTime="" lastModifiedTime="" pageLevel="1" />
	<one:Page ID="{}" name="This is another quick note" dateTime="" lastModifiedTime="" pageLevel="1" />
	</one:Section>
</one:UnfiledNotes>
*/