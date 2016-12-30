package util

import "time"

func TimeToExcelTime(t time.Time) float64 {
	return float64(t.Unix())/86400.0 + 25569.0
}
