package xlsx

import (
	"strconv"
)

type comment struct {
	Ref    string
	Author string
	Text   string
}

func newComments(commentsXml xlsxComments) []comment {
	comments := make([]comment, len(commentsXml.CommentList))
	for i, c := range commentsXml.CommentList {
		authrID, err := strconv.Atoi(c.AuthorID)
		if err == nil && authrID < len(commentsXml.Authors) {
			comments[i].Author = commentsXml.Authors[authrID]
		}
		comments[i].Text = c.Value
		comments[i].Ref = c.Ref
	}
	return comments
}
