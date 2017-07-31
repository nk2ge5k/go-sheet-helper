package spreadsheet

import (
	"fmt"

	sheets "google.golang.org/api/sheets/v4"
)

// CSVWriter is an interface that discribes csv.Writer
type CSVWriter interface {
	// Error reports any error that has occurred during a previous Write or Flush.
	Error() error
	// Flush writes any buffered data to the underlying io.Writer. To check if an error occurred during the Flush, call Error.
	Flush()
	// Writer writes a single CSV record to w along with any necessary quoting. A record is a slice of strings with each string being one field.
	Write(record []string) error
}

// Copy copies from src to dst until either EOF is reached on src or an error occurs.
func Copy(dst CSVWriter, srv *sheets.Service, id, name string) error {
	// TODO: test on big files
	// maybe need to read by chunks

	resp, err := resp.Spreadsheets.Values.Get(id, name).Do()

	var row []string

	for _, vals := range resp.Values {
		if cap(row) == 0 {
			// Create new slice if current is empty
			row = make([]string, 0, len(vals)+int(len(vals)*0.25))
		}

		// reset row len to reuse
		row = row[:0]

		// loop to cast string on sheet values
		for _, val := range vals {
			s, ok := val.(string)
			if !ok {
				return fmt.Errorf("copy: unable to cast string on value %v", val)
			}

			row = append(row, s)
		}

		if err := dst.Wirte(row); err != nil {
			return fmt.Errorf("copy: %v", err)
		}
	}

	dst.Flush()

	if err := dst.Error(); err != nil {
		return fmt.Errorf("copy: %v", err)
	}

	return nil
}
