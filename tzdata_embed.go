package main

// Embed IANA time zone database into the binary so time.LoadLocation works
// on machines without Go or system zoneinfo installed.
import _ "time/tzdata"
