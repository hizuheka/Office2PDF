package main

import (
	"testing"
)

func TestGetFilePaths(t *testing.T) {
	folderPath := "./testdata"

	expectedFilePaths := []string{
		"testdata\\test1.xlsx",
		"testdata\\test2.xlsx",
	}

	xlsPaths, _, _, err := getFilePaths(folderPath)
	if err != nil {
		t.Errorf("Error occurred while getting file paths: %v", err)
	}

	if len(xlsPaths) != len(expectedFilePaths) {
		t.Errorf("Number of file paths is incorrect. Expected: %d, Actual: %d", len(expectedFilePaths), len(xlsPaths))
	}

	for i, path := range xlsPaths {
		if path != expectedFilePaths[i] {
			t.Errorf("File path is incorrect. Expected: %s, Actual: %s", expectedFilePaths[i], path)
		}
	}
}
