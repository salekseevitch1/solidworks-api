﻿using System;

namespace CADBooster.SolidDna
{
    /// <summary>
    /// A command item for the top command bar in SolidWorks
    /// </summary>
    public class CommandManagerItem : ICommandManagerItem
    {
        /// <summary>
        /// True if the command should be added to the tab
        /// </summary>
        public bool AddToTab { get; set; } = true;

        /// <summary>
        /// The command ID of this item (set by the parent)
        /// </summary>
        public int CommandId { get; set; }

        /// <summary>
        /// The unique ID used for identifying a callback that should be associated with this item
        /// </summary>
        public string CallbackId { get; } = Guid.NewGuid().ToString("N");

        /// <summary>
        /// The help text for this item. Is only used for toolbar items and flyouts, underneath the <see cref="Tooltip"/>. Is not used for menus and flyout items.
        /// </summary>
        public string Hint { get; set; }

        /// <summary>
        /// The index position of the image to use for this item from the parent image list (zero-index)
        /// </summary>
        public int ImageIndex { get; set; }

        /// <summary>
        /// The type of item this is, such as a menu item or a toolbar item or both
        /// </summary>
        public CommandItemType ItemType { get; set; } = CommandItemType.MenuItem | CommandItemType.ToolbarItem;

        /// <summary>
        /// The name of the item. Is used as the only text for menus and flyout items. Is not used for toolbar items.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// The action to call when the item is clicked
        /// </summary>
        public Action OnClick { get; set; }

        /// <summary>
        /// The position of the item in the list. Specify 0 to add the item to the beginning of the toolbar or menu, or specify -1 to add it to the end.
        /// After creating the item, we set this to the actual position.
        /// </summary>
        public int Position { get; set; } = -1;

        /// <summary>
        /// The tab view style (whether and how to show in the large icon tab bar view)
        /// </summary>
        public CommandManagerItemTabView TabView { get; set; } = CommandManagerItemTabView.IconWithTextBelow;

        /// <summary>
        /// The name of this item. Is only used for toolbar items and flyouts, above the <see cref="Hint"/>. Is not used for menus and flyout items.
        /// </summary>
        public string Tooltip { get; set; }

        /// <summary>
        /// True to show this item in the command tab when an assembly is open. Only works for toolbar items and flyouts, not for menus or flyout items.
        /// </summary>
        public bool VisibleForAssemblies { get; set; } = true;

        /// <summary>
        /// True to show this item in the command tab when a drawing is open. Only works for toolbar items and flyouts, not for menus or flyout items.
        /// </summary>
        public bool VisibleForDrawings { get; set; } = true;

        /// <summary>
        /// True to show this item in the command tab when a part is open. Only works for toolbar items and flyouts, not for menus or flyout items.
        /// </summary>
        public bool VisibleForParts { get; set; } = true;

        /// <summary>
        /// Returns a user-friendly string with group properties.
        /// </summary>
        /// <returns></returns>
        public override string ToString() => $"Item with name: {Name}. CommandID: {CommandId}. Position: {Position}. Image index: {ImageIndex}. Hint: {Hint}. Tooltip: {Tooltip}.";
    }
}
