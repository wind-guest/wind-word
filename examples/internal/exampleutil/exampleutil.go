package exampleutil

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"runtime"
)

func OutputDir() (string, error) {
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		return "", fmt.Errorf("failed to resolve example path")
	}

	outputDir := filepath.Clean(filepath.Join(filepath.Dir(file), "..", "..", "output"))
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		return "", err
	}

	return outputDir, nil
}

func OutputPath(name string) (string, error) {
	outputDir, err := OutputDir()
	if err != nil {
		return "", err
	}

	return filepath.Join(outputDir, name), nil
}

func SamplePNGData(width, height int) ([]byte, error) {
	if width <= 0 || height <= 0 {
		return nil, fmt.Errorf("invalid image size: %dx%d", width, height)
	}

	img := image.NewRGBA(image.Rect(0, 0, width, height))
	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			r := uint8(30 + (190*x)/width)
			g := uint8(60 + (120*y)/height)
			b := uint8(120 + (80*(x+y))/(width+height))
			img.Set(x, y, color.RGBA{R: r, G: g, B: b, A: 255})
		}
	}

	for x := 0; x < width; x++ {
		img.Set(x, 0, color.RGBA{R: 255, G: 255, B: 255, A: 255})
		img.Set(x, height-1, color.RGBA{R: 255, G: 255, B: 255, A: 255})
	}
	for y := 0; y < height; y++ {
		img.Set(0, y, color.RGBA{R: 255, G: 255, B: 255, A: 255})
		img.Set(width-1, y, color.RGBA{R: 255, G: 255, B: 255, A: 255})
	}

	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}
