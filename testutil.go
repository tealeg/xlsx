package xlsx

import (
	"os"
	"path/filepath"

	qt "github.com/frankban/quicktest"
)

// cleanTempDir removes all the temporary files from NewDiskVCellStore
func cleanTempDir(c *qt.C) {
	tempDirBase := os.TempDir()

	globPattern := tempDirBase + "/" + cellStorePrefix + "*"

	dirs, err := filepath.Glob(globPattern)
	if err != nil {
		c.Logf("Cannot glob files of %s", globPattern)
		c.FailNow()
	}

	for _, directory := range dirs {
		if err = os.RemoveAll(directory); err != nil {
			c.Logf("Cannot remove files of %s", directory)
			c.FailNow()
		}
	}
}

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

	if !c.Failed() {
		cleanTempDir(c)
	}
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

	if !c.Failed() {
		cleanTempDir(c)
	}

}
