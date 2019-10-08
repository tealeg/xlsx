package xlsx

import (
	"testing"

	qt "github.com/frankban/quicktest"
)

func TestRemoveDimensionTag(t *testing.T) {
	c := qt.New(t)
	out := removeDimensionTag(`<foo><dimension ref="A1:Z20"></dimension></foo>`)
	c.Assert("<foo></foo>", qt.Equals, out)

}
