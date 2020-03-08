package xlsx

import qt "github.com/frankban/quicktest"

// csRunC will run the given test function with all available
// CellStoreConstructors.  You must take care of setting the
// CellStoreConstructors on the File struct or whereever else it is needed.
func csRunC(c *qt.C, description string, test func(c *qt.C, constructor CellStoreConstructor)) {

	c.Run(description, func(c *qt.C) {
		c.Run("MemoryCellStore", func(c *qt.C) {
			test(c, NewMemoryCellStore)
		})
		c.Run("DiskVCellStore", func(c *qt.C) {
			test(c, NewDiskVCellStore)
		})
	})
}

// csRunO will run the given test function with all available CellStore FileOptions, you must takes care of passing the FileOption to the appropriate method.
func csRunO(c *qt.C, description string, test func(c *qt.C, option FileOption)) {
	c.Run(description, func(c *qt.C) {
		c.Run("MemoryCellStore", func(c *qt.C) {
			test(c, UseMemoryCellStore)
		})
		c.Run("DiskVCellStore", func(c *qt.C) {
			test(c, UseDiskVCellStore)
		})
	})
}
