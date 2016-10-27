using System;
//using Extensibility;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Collections;
using System.Text.RegularExpressions;
using System.Linq;
using System.Linq.Expressions;

using EnvDTE80;
using System.Collections.Generic;

namespace ASEExpertVS2005
{
	/// <summary>
	/// Summary description for ASListView.
	/// </summary>
	public class ASListView : ListView
	{

		private List<ListItemData> _list = null;
		private List<ListItemData> _activeList = null;

		public void SetList(List<ListItemData> list)
		{
			_list = list;
			_activeList = new List<ListItemData>();
			_activeList.AddRange(_list);

			refresh();
		}

		public void SetFilter(string filter)
		{
			try
			{
                if (filter == "")
                {
                    _activeList.AddRange(_list);
                    return;
                }

                _activeList.Clear();

                var f = filter.ToLower().Split(' ');
                _activeList = _list.Where(x => f.All(x.Name.ToLower().Contains)).ToList();
			}
			catch
			{
			}

			refresh();
		}

		private void refresh()
		{
			Items.Clear();
            if (_activeList == null)
                return;

			foreach(ListItemData data in _activeList)
			{
				ListViewItem li = new ListViewItem();

				li.Text = data.Name;
				li.ImageIndex = data.Image;
                foreach(object item in data.Datas)
				    li.SubItems.Add(item.ToString());
                li.Tag = data;

				Items.Add(li);
			}

			if (Items.Count > 0)
				Items[0].Selected = true;
		}

		public ListItemData SelectedData
		{
			get
			{
				if (SelectedItems == null) return null;
				if (SelectedItems.Count == 0) return null;

				return (ListItemData) SelectedItems[0].Tag;
			}
		}

		public void ProcessKey(Keys key)
		{
			if (this.Items.Count < 1)
				return;
			if (this.SelectedItems == null) 
				this.Items[0].Selected = true;

			int idx = this.SelectedItems[0].Index;

			this.SelectedItems.Clear();

			if (key == Keys.Up)			OnKeyUp(idx);
			if (key == Keys.Down)		OnKeyDown(idx);
			if (key == Keys.PageUp)		OnKeyPageUp(idx);
			if (key == Keys.PageDown)	OnKeyPageDown(idx);
			if (key == Keys.Home)		OnKeyHome(idx);
			if (key == Keys.End)		OnKeyEnd(idx);

			if (this.SelectedItems == null) 
				return;
			if (this.SelectedItems.Count == 0) 
				return;

			this.SelectedItems[0].Focused = true;
			this.SelectedItems[0].EnsureVisible();
		}

		private int CountVisibleItems
		{
			get
			{
				Rectangle r = this.GetItemRect(0);
				int res = (this.Height / r.Height) - 3;
			
				return res;
			}
		}

		private void OnKeyUp(int idx)
		{
			if (idx == 0) 
				this.Items[this.Items.Count-1].Selected = true;
			else
				this.Items[idx-1].Selected = true;
		}

		private void OnKeyDown(int idx)
		{
			if (idx == this.Items.Count-1)
				this.Items[0].Selected = true;
			else
				this.Items[idx+1].Selected = true;
		}

		private void OnKeyPageUp(int idx)
		{
			this.SelectedItems.Clear();
			if ((idx ) - CountVisibleItems <= 0)
				this.Items[0].Selected = true;
			else
				this.Items[idx - CountVisibleItems].Selected = true;
		}

		private void OnKeyPageDown(int idx)
		{
			this.SelectedItems.Clear();
			if ((idx + CountVisibleItems) >= (this.Items.Count - 1))
				this.Items[this.Items.Count - 1].Selected = true;
			else
				this.Items[idx + CountVisibleItems].Selected = true;
		}

		private void OnKeyHome(int idx)
		{
			this.SelectedItems.Clear();
			this.Items[0].Selected = true;
		}

		private void OnKeyEnd(int idx)
		{
			this.SelectedItems.Clear();
			this.Items[this.Items.Count-1].Selected = true;
		}
	}
}
