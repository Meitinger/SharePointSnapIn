/* Copyright (C) 2014, Manuel Meitinger
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 2 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using SharePointSnapIn.Properties;

namespace KMOAPICapture
{
    /// <summary>
    /// Represents a list field that overcomes most OpenAPI limitations.
    /// </summary>
    public class ListFieldEx : ListField
    {
        /// <summary>
        /// Represents the virtual <see cref="ListItem"/> collection.
        /// </summary>
        public abstract class ItemCollection : Collection<ListItem>
        {
            private readonly ListItemCollection listItems;

            internal ItemCollection(ListItemCollection listItems)
            {
                // set the list and create the empty element
                this.listItems = listItems;
                listItems.Add(new ListItem(string.Empty, Resources.EmptySelection, true));
            }

            /// <summary>
            /// This is an infrastructure method and shouldn't be used directly.
            /// </summary>
            protected sealed override void ClearItems()
            {
                // clear the items and re-add the empty element
                base.ClearItems();
                listItems.Clear();
                listItems.Add(new ListItem(string.Empty, Resources.EmptySelection, true));
            }

            /// <summary>
            /// This is an infrastructure method and shouldn't be used directly.
            /// </summary>
            protected sealed override void InsertItem(int index, ListItem item)
            {
                // insert the item
                base.InsertItem(index, item);
                listItems.Insert(index + 1, TranslateItem(item));
                Update();
            }

            /// <summary>
            /// This is an infrastructure method and shouldn't be used directly.
            /// </summary>
            protected sealed override void RemoveItem(int index)
            {
                // remove the item
                base.RemoveItem(index);
                listItems.RemoveAt(index + 1);
                Update();
            }

            /// <summary>
            /// This is an infrastructure method and shouldn't be used directly.
            /// </summary>
            protected sealed override void SetItem(int index, ListItem item)
            {
                // change the item
                base.SetItem(index, item);
                listItems[index + 1] = TranslateItem(item);
                Update();
            }

            /// <summary>
            /// Refreshes the underlying <see cref="ListField"/> and actual <see cref="ListItem"/> after a property of the virtual <see cref="ListItem"/> at the given position was changed.
            /// </summary>
            /// <param name="index">The position of the changed <see cref="ListItem"/>.</param>
            public void Refresh(int index)
            {
                // check the bounds and reset the item
                if (index < 0 || index >= Items.Count)
                    throw new ArgumentOutOfRangeException("index");
                listItems[index + 1] = TranslateItem(Items[index]);
                Update();
            }

            /// <summary>
            /// Refreshes the underlying <see cref="ListField"/> and all actual <see cref="ListItem"/>s.
            /// </summary>
            public void RefreshAll()
            {
                // update all items
                for (var index = 0; index < Items.Count; index++)
                    listItems[index + 1] = TranslateItem(Items[index]);
                Update();
            }

            /// <summary>
            /// This is an infrastructure method and shouldn't be used directly.
            /// </summary>
            protected IEnumerable<int> GetSelectedIndices()
            {
                // get all selected list item indices
                for (var index = 0; index < Items.Count; index++)
                    if (listItems[index + 1].Selected)
                        yield return index;
            }

            /// <summary>
            /// This is an infrastructure property and shouldn't be used directly.
            /// </summary>
            protected ListItem EmptyItem { get { return listItems[0]; } }

            /// <summary>
            /// This is an infrastructure method and shouldn't be used directly.
            /// </summary>
            protected abstract ListItem TranslateItem(ListItem item);

            /// <summary>
            /// This is an infrastructure method and shouldn't be used directly.
            /// </summary>
            protected abstract void Update();

            internal abstract void HandleValueIdChange();
        }

        private class MultiListItemCollection : ItemCollection
        {
            internal MultiListItemCollection(ListItemCollection listItems) : base(listItems) { }

            protected override ListItem TranslateItem(ListItem item)
            {
                // return a new selection item based on the list item
                return new ListItem(string.Empty, string.Format(item.Selected ? Resources.ExcludeSelectionFormat : Resources.IncludeSelectionFormat, item.Text), false);
            }

            protected override void Update()
            {
                // update and select the empty item
                var selected = Items.Where(e => e.Selected);
                EmptyItem.Text = !selected.Any() ? Resources.EmptySelection : string.Join(", ", selected.Select(e => e.Text));
                EmptyItem.Selected = true;
            }

            internal override void HandleValueIdChange()
            {
                // toggle the selection
                foreach (var index in GetSelectedIndices())
                {
                    var item = Items[index];
                    item.Selected = !item.Selected;
                }

                // refresh all items
                RefreshAll();
            }
        }

        private class SingleListItemCollection : ItemCollection
        {
            internal SingleListItemCollection(ListItemCollection listItems) : base(listItems) { }

            protected override ListItem TranslateItem(ListItem item)
            {
                // return the item as-is
                return item;
            }

            protected override void Update()
            {
                // select the empty item if nothing is selected
                EmptyItem.Selected = !Items.Any(e => e.Selected);
            }

            internal override void HandleValueIdChange()
            {
                // simply update
                Update();
            }
        }

        /// <summary>
        /// Creates a new advances <see cref="ListField"/>.
        /// </summary>
        public ListFieldEx()
            : base()
        {
            this.Init();
        }

        private ListFieldEx(BaseField baseField)
            : base(baseField)
        {
            this.Init();
        }

        private void Init()
        {
            // disable multiselect at the base and create the items
            base.AllowMultipleSelection = false;
            this.Items = new SingleListItemCollection(base.Items);
        }

        /// <summary>
        /// Gets or sets whether mutliple <see cref="ListItem"/>s can be selected.
        /// </summary>
        public new bool AllowMultipleSelection
        {
            get { return this.Items is MultiListItemCollection; }
            set
            {
                if (value != this.AllowMultipleSelection)
                {
                    // make a copy of the current items 
                    var items = this.Items.ToArray();

                    // clear the base items and set the new value
                    base.Items.Clear();
                    this.Items = value ? (ItemCollection)new MultiListItemCollection(base.Items) : (ItemCollection)new SingleListItemCollection(base.Items);

                    // re-add the items
                    foreach (var item in items)
                        this.Items.Add(item);
                }
            }
        }

        /// <summary>
        /// Clones the current <see cref="ListFieldEx"/>.
        /// </summary>
        /// <returns>A copy of the field with the same settings and <see cref="ListItem"/>s</returns>
        public override object Clone()
        {
            // clone the field
            var listField = new ListFieldEx(this)
            {
                RaiseFindEvent = this.RaiseFindEvent,
                AllowMultipleSelection = this.AllowMultipleSelection,
                ButtonSize = this.ButtonSize,
            };
            foreach (var item in this.Items)
                listField.Items.Add(new ListItem(item.Value, item.Text, item.Selected));
            return listField;
        }

        /// <summary>
        /// Gets the virtual <see cref="ListItem"/> collection.
        /// </summary>
        public new ItemCollection Items { get; private set; }

        /// <summary>
        /// Gets or sets the comma-separated value of all selected <see cref="ListItem"/>s.
        /// </summary>
        /// <exception cref="ArgumentNullException">The value is <c>null</c>.</exception>
        /// <exception cref="ArgumentOutOfRangeException">The value contains multiple values but <see cref="AllowMultipleSelection"/> is <c>false</c>.</exception>
        public override string Value
        {
            get { return string.Join(",", this.Items.Where(e => e.Selected).Select(e => e.Value)); }
            set
            {
                // check and split the value
                if (value == null)
                    throw new ArgumentNullException("Value");
                var values = new HashSet<string>(value.Split(','));

                // make sure multiple selections are allowed if needed
                if (values.Count > 1 && !this.AllowMultipleSelection)
                    throw new ArgumentOutOfRangeException("Value");

                // set the selections and update the items
                foreach (var item in this.Items)
                    item.Selected = values.Contains(item.Value);
                this.Items.RefreshAll();
            }
        }

        /// <summary>
        /// This is an infrastructure property and shouldn't be used directly.
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public override string ValueId
        {
            get { return base.ValueId; }
            set
            {
                // perform the actual operation and resync the items
                base.ValueId = value;
                this.Items.HandleValueIdChange();
            }
        }
    }
}
