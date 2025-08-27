package main

import (
	"errors"
	"flag"
	"fmt"
	"log"
	"math/rand"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// ==================== Types ====================

type RoleMap struct {
	Role         string
	SourceColumn string
	Service      string // "07" | "10" | "both"
	Slots07      int
	Slots10      int
}

type Person struct {
	Name      string
	IsPenatua bool
	Marks     map[string]bool // normalized header -> eligible
}

type Assignment = map[time.Time]map[string]map[string][]string // date -> service -> role -> names

// ==================== Flags ====================

var (
	bulanFlag   = flag.String("bulan", "", "Bulan (1-12 atau nama Indonesia, wajib)")
	tahunFlag   = flag.Int("tahun", 0, "Tahun (wajib)")
	tanggalFlag = flag.Int("tgl", 0, "Tanggal (opsional)")

	maxLektorFlag = flag.Int("maxLektor", 2, "Jumlah Lektor per ibadah (default 2, maks 4)")
	maxProkantor  = flag.Int("maxProkantor", 2, "Jumlah Prokantor (default 2, maks 3)")
	maxPemusik    = flag.Int("maxPemusik", 2, "Jumlah Pemusik (default 2, maks 3)")

	seedFlag     = flag.Int64("seed", 0, "Seed RNG (opsional, 0=acak)")
	outdirFlag   = flag.String("outdir", "", "Folder output")
	templateName = flag.String("template", "TemplateOutput.xlsx", "Nama template")

	// Tambahan: jumlah baris header yang discan placeholder-nya
	headerRowsFlag = flag.Int("headerRows", 30, "Jumlah baris atas untuk scan placeholder header (default 30)")
	masterOverride  = flag.String("master", "", "Path Master.xlsx khusus")
	forceMasterCopy = flag.Bool("forceMasterCopy", false, "Paksa salin Master.xlsx")

	verboseFlag = flag.Bool("v", false, "Verbose mode")

	kolektanPatternFlag = flag.String("kolektanPattern", "2b", "Pola Kolektan (1a..4e)")
	pJemaatPatternFlag  = flag.String("pjemaatPattern", "3a", "Pola P. Jemaat (1a..4e)")

	// Hardening flags
	strictCompositionFlag = flag.Bool("strictComposition", false, "Strict komposisi P/J: bila kuota tidak tercapai, sisanya kosong (tanpa relax-any)")
	noRelaxB2BFlag        = flag.Bool("noRelaxB2B", false, "Nonaktifkan relax back-to-back (prefer anti-B2B wajib dipatuhi)")
)

func main() {
	log.SetFlags(0)
	flag.Parse()
	if err := run(); err != nil {
		fmt.Fprintln(os.Stderr, "ERROR:", err)
		os.Exit(1)
	}
}

func isVerbose() bool { return *verboseFlag }

// ==================== run() ====================

func run() error {
	// RNG
	if *seedFlag != 0 {
		rand.Seed(*seedFlag)
	} else {
		rand.Seed(time.Now().UnixNano())
	}
	if *bulanFlag == "" || *tahunFlag == 0 {
		return errors.New("parameter -bulan dan -tahun wajib; contoh: -bulan Agustus -tahun 2025")
	}
	month, err := parseMonth(*bulanFlag)
	if err != nil {
		return err
	}
	year := *tahunFlag

	// Ensure config dir & Master.xlsx
	docDir := getDocumentsDir()
	baseDir := filepath.Join(docDir, "JadwalPetugas")
	configDir := filepath.Join(baseDir, "config")
	if err := os.MkdirAll(configDir, 0o755); err != nil {
		return fmt.Errorf("membuat folder %s: %w", configDir, err)
	}
	exedir, _ := exeDir()
	cwd, _ := os.Getwd()

	var masterPath string
	if s := strings.TrimSpace(*masterOverride); s != "" {
		masterPath = s
	} else {
		masterAtConfig := filepath.Join(configDir, "Master.xlsx")
		candidates := []string{filepath.Join(cwd, "Master.xlsx"), filepath.Join(exedir, "Master.xlsx")}
		var src string
		for _, c := range candidates {
			if _, err := os.Stat(c); err == nil {
				src = c
				break
			}
		}
		if *forceMasterCopy {
			if src == "" {
				return fmt.Errorf("Master.xlsx sumber tidak ditemukan")
			}
			if err := copyFile(src, masterAtConfig); err != nil {
				return err
			}
			if isVerbose() {
				fmt.Println("Master.xlsx ditimpa dari", src, "->", masterAtConfig)
			}
		} else {
			if _, err := os.Stat(masterAtConfig); os.IsNotExist(err) {
				if src == "" {
					return fmt.Errorf("Master.xlsx tidak ditemukan")
				}
				if err := copyFile(src, masterAtConfig); err != nil {
					return err
				}
				if isVerbose() {
					fmt.Println("Master.xlsx disalin ke", masterAtConfig, "dari", src)
				}
			}
		}
		masterPath = masterAtConfig
	}

	people, mappings, err := loadMaster(masterPath)
	if err != nil {
		return fmt.Errorf("memuat Master.xlsx: %w", err)
	}
	if len(people) == 0 {
		return errors.New("Sheet Petugas kosong/invalid")
	}
	if len(mappings) == 0 {
		return errors.New("Sheet MappingRole kosong/invalid")
	}

	loc := mustLoc("Asia/Jakarta")
	var dates []time.Time
	if *tanggalFlag > 0 {
		d, err := safeDate(year, month, *tanggalFlag, loc)
		if err != nil {
			return err
		}
		dates = []time.Time{d}
	} else {
		dates = allSundays(year, month, loc)
		if len(dates) == 0 {
			return errors.New("tidak ada hari Minggu pada bulan ini")
		}
	}

	maxLektor := clamp(*maxLektorFlag, 1, 4)
	maxPro := clamp(*maxProkantor, 1, 3)
	maxMus := clamp(*maxPemusik, 1, 3)

	kPen, kJem, _, err := parsePattern(*kolektanPatternFlag)
	if err != nil {
		return fmt.Errorf("pola Kolektan: %w", err)
	}
	pPen, pJem, _, err := parsePattern(*pJemaatPatternFlag)
	if err != nil {
		return fmt.Errorf("pola P. Jemaat: %w", err)
	}

	if isVerbose() {
		fmt.Printf("Flags: strictComposition=%v, noRelaxB2B=%v, seed=%d\n", *strictCompositionFlag, *noRelaxB2BFlag, *seedFlag)
		fmt.Printf("Limits: Lektor=%d Prokantor=%d Pemusik=%d\n", maxLektor, maxPro, maxMus)
		fmt.Printf("HeaderRows: %d\n", *headerRowsFlag)
		fmt.Printf("Pattern: Kolektan=%s (P:%d J:%d) | P.Jemaat=%s (P:%d J:%d)\n",
			*kolektanPatternFlag, kPen, kJem, *pJemaatPatternFlag, pPen, pJem)
	}

	assign := make(Assignment)
	if err := generate(assign, dates, people, mappings, maxLektor, maxPro, maxMus, loc, isVerbose(), kPen, kJem, pPen, pJem); err != nil {
		return err
	}

	// Output
	outDir := *outdirFlag
	if strings.TrimSpace(outDir) == "" {
		outDir = baseDir
	}
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		return err
	}
	now := time.Now().In(loc)
	outName := fmt.Sprintf("JadwalPetugas_%s_%02d.%02d.%02d.xlsx", monthNameID(month), now.Hour(), now.Minute(), now.Second())
	outPath := filepath.Join(outDir, outName)

	if err := writeTemplateAware(assign, mappings, dates, exedir, *templateName, outPath, loc, isVerbose()); err != nil {
		return err
	}
	fmt.Println("SUKSES:", outPath)
	return nil
}

// ==================== loadMaster() ====================

func loadMaster(path string) ([]Person, []RoleMap, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, nil, err
	}
	defer f.Close()

	petugasSheet := findSheet(f, []string{"Petugas"})
	if petugasSheet == "" {
		return nil, nil, errors.New("Sheet Petugas tidak ditemukan")
	}
	mappingSheet := findSheet(f, []string{"MappingRole"})
	if mappingSheet == "" {
		return nil, nil, errors.New("Sheet MappingRole tidak ditemukan")
	}

	petRows, _ := f.GetRows(petugasSheet)
	if len(petRows) < 2 {
		return nil, nil, errors.New("Petugas kosong")
	}

	// Header index
	headIdx := map[string]int{}
	for col, name := range petRows[0] {
		headIdx[strings.ToLower(strings.TrimSpace(name))] = col
	}
	nameCol, ok := headIdx["nama"]
	if !ok {
		return nil, nil, errors.New("Kolom Nama wajib")
	}
	penatuaCol := -1
	if idx, ok := headIdx["penatua"]; ok {
		penatuaCol = idx
	}

	var people []Person
	for i := 1; i < len(petRows); i++ {
		row := petRows[i]
		if nameCol >= len(row) {
			continue
		}
		name := strings.TrimSpace(row[nameCol])
		if name == "" {
			continue
		}
		p := Person{Name: name, Marks: map[string]bool{}}
		if penatuaCol >= 0 && penatuaCol < len(row) {
			p.IsPenatua = isMarked(row[penatuaCol])
		}
		for k, v := range row {
			if k >= len(petRows[0]) {
				continue
			}
			hdr := strings.TrimSpace(petRows[0][k])
			if hdr == "" {
				continue
			}
			p.Marks[normKey(hdr)] = isMarked(v)
		}
		people = append(people, p)
	}

	relRows, _ := f.GetRows(mappingSheet)
	if len(relRows) < 2 {
		return people, nil, errors.New("Mapping kosong")
	}
	mh := indexHeader(relRows[0])
	roleCol := findHeader(mh, []string{"role"})
	srcCol := findHeader(mh, []string{"kolom master", "source"})
	serviceCol := findHeader(mh, []string{"service"})
	slots07Col := findHeader(mh, []string{"slots07"})
	slots10Col := findHeader(mh, []string{"slots10"})
	if roleCol < 0 || srcCol < 0 {
		return people, nil, errors.New("MappingRole wajib ada kolom Role & Kolom Master")
	}

	var maps []RoleMap
	for i := 1; i < len(relRows); i++ {
		row := relRows[i]
		if roleCol >= len(row) || srcCol >= len(row) {
			continue
		}
		role := strings.TrimSpace(row[roleCol])
		src := strings.TrimSpace(row[srcCol])
		if role == "" || src == "" {
			continue
		}
		m := RoleMap{Role: role, SourceColumn: src, Service: "both"}
		if serviceCol >= 0 && serviceCol < len(row) {
			v := strings.TrimSpace(strings.ToLower(row[serviceCol]))
			if v == "07" {
				m.Service = "07"
			}
			if v == "10" {
				m.Service = "10"
			}
		}
		if slots07Col >= 0 && slots07Col < len(row) {
			m.Slots07 = atoiSafe(row[slots07Col])
		}
		if slots10Col >= 0 && slots10Col < len(row) {
			m.Slots10 = atoiSafe(row[slots10Col])
		}
		maps = append(maps, m)
	}
	return people, maps, nil
}

// ==================== generate() ====================

func generate(assign Assignment, dates []time.Time, people []Person, maps []RoleMap,
	maxLektor, maxPro, maxMus int, loc *time.Location, verbose bool,
	kolektanPen, kolektanJem, pjemaatPen, pjemaatJem int) error {

	lastAssigned := map[string]time.Time{}

	// index Penatua untuk rekap cepat
	penIdx := map[string]bool{}
	for _, p := range people {
		penIdx[p.Name] = p.IsPenatua
	}

	for di, d := range dates {
		if assign[d] == nil {
			assign[d] = map[string]map[string][]string{}
		}
		services := []string{"07", "10"}
		assigned07 := map[string]bool{}
		assigned10 := map[string]bool{}
		assignedAnyToday := map[string]bool{}

		if verbose {
			fmt.Printf("=== %s ===\n", d.Format("Mon, 02 Jan 2006"))
		}

		for _, svc := range services {
			if assign[d][svc] == nil {
				assign[d][svc] = map[string][]string{}
			}
			if verbose {
				fmt.Printf("  [Service %s]\n", svc)
			}

			// one-line summary holders untuk komposisi
			compStatus := map[string]string{"kolektan": "N/A", "pjemaat": "N/A"}

			grouped, others := groupMappingsForService(maps, svc)

			// ---- Split others menjadi MP vs non-MP
			mpRows := []RoleMap{}
			otherNonMP := []RoleMap{}
			for _, m := range others {
				if m.Service != "both" && m.Service != svc {
					continue
				}
				if isMajelisPendamping(m.Role) {
					mpRows = append(mpRows, m)
				} else {
					otherNonMP = append(otherNonMP, m)
				}
			}

			// ---- prefer function (hindari back-to-back Minggu berurutan)
			var prevSunday time.Time
			if di > 0 {
				prevSunday = dates[di-1]
			}
			prefer := func(name string) bool {
				if prevSunday.IsZero() {
					return true
				}
				if t, ok := lastAssigned[name]; ok && sameDay(t, prevSunday) {
					return false
				}
				return true
			}

			// ======================================================
			// 1) Majelis Pendamping (prioritas pertama, hanya 10.00)
			// ======================================================
			if svc == "10" && len(mpRows) > 0 {
				for _, m := range mpRows {
					slots := 1
					if m.Slots10 > 0 {
						slots = m.Slots10
					}
					cands := filterCandidates(people, m.SourceColumn, true) // wajib Penatua
					rand.Shuffle(len(cands), func(i, j int) { cands[i], cands[j] = cands[j], cands[i] })

					picked := []string{}
					// (a) hormati prefer (hindari back-to-back), no double-role 10.00, no multi-role/day
					for _, name := range cands {
						if len(picked) >= slots {
							break
						}
						if assigned10[name] || assignedAnyToday[name] {
							continue
						}
						if prefer(name) {
							picked = append(picked, name)
							assigned10[name] = true
							assignedAnyToday[name] = true
							lastAssigned[name] = d
						}
					}
					// (b) RELAX khusus MP: boleh ambil dari yang sudah bertugas 07.00 hari sama
					if len(picked) < slots {
						for _, name := range cands {
							if len(picked) >= slots {
								break
							}
							if assigned10[name] {
								continue // tetap jangan dua peran di 10.00
							}
							// izinkan meski assignedAnyToday[name] == true (dari 07.00)
							picked = append(picked, name)
							assigned10[name] = true
							assignedAnyToday[name] = true
							lastAssigned[name] = d
							if verbose {
								fmt.Printf("      pick(MP-relax) %-20s\n", name)
							}
						}
					}
					assign[d][svc][m.Role] = picked
				}
			}

			// ======================================================
			// 2) Komposisi: Kolektan & P. Jemaat (kedua)
			// ======================================================
			for _, key := range []string{"kolektan", "pjemaat"} {
				rows := grouped[key]
				if len(rows) == 0 {
					continue
				}
				var needPen, needJem int
				if key == "kolektan" {
					needPen, needJem = kolektanPen, kolektanJem
				}
				if key == "pjemaat" {
					needPen, needJem = pjemaatPen, pjemaatJem
				}

				totalNeed := needPen + needJem
				if totalNeed > len(rows) {
					totalNeed = len(rows)
				}

				penNames, jemNames := []string{}, []string{}
				for _, rm := range rows {
					p, j := filterCandidatesSplit(people, rm.SourceColumn)
					penNames = append(penNames, p...)
					jemNames = append(jemNames, j...)
				}
				penNames = uniq(penNames)
				jemNames = uniq(jemNames)
				if verbose {
					fmt.Printf("    %s pool => penatua:%d, jemaat:%d (need P:%d J:%d)\n",
						key, len(penNames), len(jemaatNames(jemNames)), needPen, needJem)
				}

				var candPen, candJem []Person
				for _, n := range penNames {
					candPen = append(candPen, Person{Name: n, IsPenatua: true})
				}
				for _, n := range jemNames {
					candJem = append(candJem, Person{Name: n, IsPenatua: false})
				}
				rand.Shuffle(len(candPen), func(i, j int) { candPen[i], candPen[j] = candPen[j], candPen[i] })
				rand.Shuffle(len(candJem), func(i, j int) { candJem[i], candJem[j] = candJem[j], candJem[i] })

				var already map[string]bool
				if svc == "07" {
					already = assigned07
				} else {
					already = assigned10
				}
				picked := pickWithComposition(candPen, candJem, needPen, needJem, prefer, already, assignedAnyToday, verbose)
				if len(picked) > totalNeed {
					picked = picked[:totalNeed]
				}
				for i, rm := range rows {
					if i < len(picked) {
						assign[d][svc][rm.Role] = []string{picked[i]}
						lastAssigned[picked[i]] = d
					} else {
						assign[d][svc][rm.Role] = []string{}
					}
				}

				// --- Summary per service untuk komposisi (display only)
				if verbose {
					// count actual P/J dari picked
					countP := 0
					for _, n := range picked {
						if penIdx[n] {
							countP++
						}
					}
					countJ := len(picked) - countP

					reqTotal := totalNeed
					rem := reqTotal
					reqP := needPen
					if reqP > rem {
						reqP = rem
					}
					rem -= reqP
					reqJ := needJem
					if reqJ > rem {
						reqJ = rem
					}

					missingP := 0
					if countP < reqP {
						missingP = reqP - countP
					}
					missingJ := 0
					if countJ < reqJ {
						missingJ = reqJ - countJ
					}
					missingSlots := reqTotal - len(picked)

					status := "OK"
					if missingP > 0 || missingJ > 0 || (*strictCompositionFlag && missingSlots > 0) {
						status = fmt.Sprintf("KURANG (P:%d J:%d slot:%d)", missingP, missingJ, missingSlots)
					}
					fmt.Printf("    Rekap komposisi %s (%s): %s\n", strings.Title(key), svc, status)
					compStatus[key] = status
					if *strictCompositionFlag && missingSlots > 0 {
						fmt.Printf("      (kosong: kuota tidak terpenuhi dengan prefer anti-B2B)\n")
					}
				}
			}

			// ======================================================
			// 3) Lektor / Prokantor / Pemusik (ketiga)
			// ======================================================
			for _, g := range []struct {
				key   string
				limit int
			}{
				{"lektor", maxLektor}, {"prokantor", maxPro}, {"pemusik", maxMus},
			} {
				rows := grouped[g.key]
				if len(rows) == 0 {
					continue
				}
				limit := g.limit
				if limit > len(rows) {
					limit = len(rows)
				}
				if verbose {
					fmt.Printf("    - Group %-10s | Rows: %d | Limit: %d\n", g.key, len(rows), limit)
				}
				src := rows[0].SourceColumn
				names := filterCandidates(people, src, false) // tidak wajib Penatua
				rand.Shuffle(len(names), func(i, j int) { names[i], names[j] = names[j], names[i] })

				var already map[string]bool
				if svc == "07" {
					already = assigned07
				} else {
					already = assigned10
				}

				picked := []string{}
				for _, name := range names {
					if len(picked) >= limit {
						break
					}
					if already[name] || assignedAnyToday[name] {
						continue
					}
					if prefer(name) {
						picked = append(picked, name)
						already[name] = true
						assignedAnyToday[name] = true
						lastAssigned[name] = d
						if verbose {
							fmt.Printf("      pick %-20s\n", name)
						}
					}
				}

				// RELAX phase (fill remaining) -> ONLY if noRelaxB2B is OFF
				if !*noRelaxB2BFlag && len(picked) < limit {
					for _, name := range names {
						if len(picked) >= limit {
							break
						}
						if already[name] || assignedAnyToday[name] {
							continue
						}
						picked = append(picked, name)
						already[name] = true
						assignedAnyToday[name] = true
						lastAssigned[name] = d
						if verbose {
							fmt.Printf("      pick(relax) %-12s\n", name)
						}
					}
				}

				for i, rm := range rows {
					if i < len(picked) {
						assign[d][svc][rm.Role] = []string{picked[i]}
					} else {
						assign[d][svc][rm.Role] = []string{}
					}
				}
			}

			// ======================================================
			// 4) Role lainnya (non-MP)
			// ======================================================
			for _, m := range otherNonMP {
				if m.Service != "both" && m.Service != svc {
					continue
				}
				if svc == "07" && isMajelisPendamping(m.Role) {
					continue // safety
				}

				slots := defaultSlotsForRole(m.Role, svc, maxLektor, maxPro, maxMus)
				if svc == "07" && m.Slots07 > 0 {
					slots = m.Slots07
				}
				if svc == "10" && m.Slots10 > 0 {
					slots = m.Slots10
				}

				cands := filterCandidates(people, m.SourceColumn, isMajelisPendamping(m.Role))
				rand.Shuffle(len(cands), func(i, j int) { cands[i], cands[j] = cands[j], cands[i] })

				var already map[string]bool
				if svc == "07" {
					already = assigned07
				} else {
					already = assigned10
				}

				picked := []string{}
				for _, name := range cands {
					if len(picked) >= slots {
						break
					}
					if already[name] || assignedAnyToday[name] {
						continue
					}
					if prefer(name) {
						picked = append(picked, name)
						already[name] = true
						assignedAnyToday[name] = true
						lastAssigned[name] = d
					}
				}
				// RELAX phase -> ONLY if noRelaxB2B is OFF
				if !*noRelaxB2BFlag && len(picked) < slots {
					for _, name := range cands {
						if len(picked) >= slots {
							break
						}
						if already[name] || assignedAnyToday[name] {
							continue
						}
						picked = append(picked, name)
						already[name] = true
						assignedAnyToday[name] = true
						lastAssigned[name] = d
					}
				}
				assign[d][svc][m.Role] = picked
			}

			// One-line summary per service (Kolektan & P. Jemaat)
			if verbose {
				fmt.Printf("    Summary %s.00: Kolektan %s | P.Jemaat %s\n", svc, compStatus["kolektan"], compStatus["pjemaat"])
			}
		}
	}
	return nil
}

// ==================== Grouping & Picker ====================

func groupMappingsForService(maps []RoleMap, svc string) (map[string][]RoleMap, []RoleMap) {
	groups := map[string][]RoleMap{}
	var others []RoleMap
	for _, m := range maps {
		if m.Service != "both" && m.Service != svc {
			continue
		}
		base := baseRole(m.Role)
		switch base {
		case "lektor", "prokantor", "pemusik", "kolektan", "pjemaat":
			groups[base] = append(groups[base], m)
		default:
			others = append(others, m)
		}
	}
	return groups, others
}

func baseRole(role string) string {
	r := strings.ToLower(strings.TrimSpace(role))
	if strings.HasPrefix(r, "lektor") {
		return "lektor"
	}
	if strings.HasPrefix(r, "prokantor") {
		return "prokantor"
	}
	if strings.HasPrefix(r, "pemusik") {
		return "pemusik"
	}
	if strings.HasPrefix(r, "kolektan") {
		return "kolektan"
	}
	if strings.Contains(r, "pjemaat") || strings.Contains(r, "p. jemaat") {
		return "pjemaat"
	}
	return r
}

func pickWithComposition(
	candPen, candJem []Person,
	needPen, needJem int,
	prefer func(string) bool,
	already map[string]bool,
	assignedAnyToday map[string]bool,
	verbose bool,
) []string {
	totalNeed := needPen + needJem
	picked := []string{}

	used := map[string]bool{}

	remaining := func(pool []Person) []Person {
		res := []Person{}
		for _, p := range pool {
			if used[p.Name] || already[p.Name] || assignedAnyToday[p.Name] {
				continue
			}
			res = append(res, p)
		}
		return res
	}

	pickFrom := func(pool []Person, need *int, usePrefer bool, tag string) {
		for _, p := range pool {
			if len(picked) >= totalNeed {
				break
			}
			if *need <= 0 {
				break
			}
			if used[p.Name] || already[p.Name] || assignedAnyToday[p.Name] {
				continue
			}
			if usePrefer && !prefer(p.Name) {
				continue
			}
			picked = append(picked, p.Name)
			used[p.Name] = true
			already[p.Name] = true
			assignedAnyToday[p.Name] = true
			*need--
			if verbose {
				if tag != "" {
					fmt.Printf("      %s %-20s\n", tag, p.Name)
				} else {
					fmt.Printf("      pick %-20s\n", p.Name)
				}
			}
		}
	}

	// Step A: penuhi kuota dengan prefer (anti back-to-back)
	pickFrom(candPen, &needPen, true, "")
	pickFrom(candJem, &needJem, true, "")

	// Step B: fallback tetap menjaga kuota per tipe (prefer masih dihormati)
	if needPen > 0 {
		pickFrom(remaining(candPen), &needPen, true, "pick(fallback-P)")
	}
	if needJem > 0 {
		pickFrom(remaining(candJem), &needJem, true, "pick(fallback-J)")
	}

	// Step C: relax back-to-back per tipe (abaikan prefer) -> ONLY if noRelaxB2B OFF
	if !*noRelaxB2BFlag {
		if needPen > 0 {
			pickFrom(remaining(candPen), &needPen, false, "pick(relax-P)")
		}
		if needJem > 0 {
			pickFrom(remaining(candJem), &needJem, false, "pick(relax-J)")
		}
	}

	// Step D: kalau masih belum penuh totalNeed, isi apa saja (hanya jika tidak strict)
	if !*strictCompositionFlag && len(picked) < totalNeed {
		merged := append(remaining(candPen), remaining(candJem)...)
		rand.Shuffle(len(merged), func(i, j int) { merged[i], merged[j] = merged[j], merged[i] })
		extra := totalNeed - len(picked)
		pickFrom(merged, &extra, false, "pick(relax-any)")
	}

	return picked
}

func filterCandidatesSplit(people []Person, src string) (penatua []string, jemaat []string) {
	key := normKey(src)
	for _, p := range people {
		if mark, ok := p.Marks[key]; ok && mark {
			if p.IsPenatua {
				penatua = append(penatua, p.Name)
			} else {
				jemaat = append(jemaat, p.Name)
			}
		}
	}
	return
}

// ==================== Writer ====================

func writeTemplateAware(assign Assignment, maps []RoleMap, dates []time.Time,
	exeDir, templateFile, outPath string, loc *time.Location, verbose bool) error {
	cwd, _ := os.Getwd()
	tplPath := filepath.Join(cwd, templateFile)
	if _, err := os.Stat(tplPath); err != nil {
		tplPath = filepath.Join(exeDir, templateFile)
	}
	if err := copyFile(tplPath, outPath); err != nil {
		return err
	}
	f, err := excelize.OpenFile(outPath)
	if err != nil {
		return err
	}
	defer f.Close()
	sheet := "Jadwal Bulanan"

	// --- Fill header placeholders per tanggal (kolom) ---
	for i, d := range dates {
		col := 2 + i // B=2
		// Cakup header 07.00 & 10.00 (default 30 baris; bisa diubah dengan -headerRows)
		for r := 1; r <= *headerRowsFlag; r++ {
			addr := cell(col, r)
			val, _ := f.GetCellValue(sheet, addr)
			if strings.Contains(val, "{") {
				newv := replacePlaceholders(val, d, loc)
				if newv != val {
					_ = f.SetCellStr(sheet, addr, newv)
				}
			}
		}
	}

	// --- Hide unused columns (assume 5 slots: B..F) ---
	totalSlots := 5
	if len(dates) < totalSlots {
		for i := len(dates); i < totalSlots; i++ {
			col := 2 + i
			colName, _ := excelize.ColumnNumberToName(col)
			_ = f.SetColVisible(sheet, colName, false)
		}
	}

	// --- Write assignment values ---
	for i, d := range dates {
		col := 2 + i
		// 07.00
		for role, vals := range assign[d]["07"] {
			row := rowForRole(f, sheet, role, true)
			if row < 1 {
				if verbose {
					fmt.Println("WARN: role", role, "tidak ditemukan di template (07.00)")
				}
				continue
			}
			_ = f.SetCellStr(sheet, cell(col, row), strings.Join(vals, "\n"))
		}
		// 10.00
		for role, vals := range assign[d]["10"] {
			row := rowForRole(f, sheet, role, false)
			if row < 1 {
				if verbose {
					fmt.Println("WARN: role", role, "tidak ditemukan di template (10.00)")
				}
				continue
			}
			_ = f.SetCellStr(sheet, cell(col, row), strings.Join(vals, "\n"))
		}
	}
	return f.Save()
}

func rowForRole(f *excelize.File, sheet, role string, umum bool) int {
	rows, _ := f.GetRows(sheet)
	target := strings.TrimSpace(role)
	// 1) exact match (case-insensitive)
	for i, r := range rows {
		if len(r) > 0 && strings.EqualFold(strings.TrimSpace(r[0]), target) {
			return i + 1
		}
	}
	// 2) fuzzy khusus Majelis Pendamping
	if isMajelisPendamping(role) {
		for i, r := range rows {
			if len(r) == 0 {
				continue
			}
			lab := strings.ToLower(strings.TrimSpace(r[0]))
			if strings.Contains(lab, "majel") && strings.Contains(lab, "pend") {
				return i + 1
			}
		}
	}
	return -1
}

// ==================== Utilities ====================

func normKey(s string) string { return strings.ToLower(strings.TrimSpace(s)) }

func exeDir() (string, error) {
	p, err := os.Executable()
	if err != nil {
		return "", err
	}
	return filepath.Dir(p), nil
}

func getDocumentsDir() string {
	home, _ := os.UserHomeDir()
	if runtime.GOOS == "windows" {
		return filepath.Join(home, "Documents")
	}
	return filepath.Join(home, "Documents")
}

func copyFile(src, dst string) error {
	b, err := os.ReadFile(src)
	if err != nil {
		return err
	}
	return os.WriteFile(dst, b, 0o644)
}

func findSheet(f *excelize.File, names []string) string {
	all := f.GetSheetList()
	for _, s := range all {
		for _, n := range names {
			if strings.EqualFold(s, n) {
				return s
			}
		}
	}
	return ""
}

func isMarked(v string) bool {
	vv := strings.TrimSpace(strings.ToLower(v))
	return vv == "x" || vv == "1" || vv == "true" || vv == "ya"
}

func indexHeader(head []string) map[string]int {
	m := map[string]int{}
	for i, h := range head {
		m[strings.ToLower(strings.TrimSpace(h))] = i
	}
	return m
}
func findHeader(idx map[string]int, cands []string) int {
	for _, c := range cands {
		if v, ok := idx[strings.ToLower(c)]; ok {
			return v
		}
	}
	return -1
}

func atoiSafe(s string) int { var x int; fmt.Sscanf(strings.TrimSpace(s), "%d", &x); return x }

func clamp(v, lo, hi int) int {
	if v < lo {
		return lo
	}
	if v > hi {
		return hi
	}
	return v
}

func mustLoc(name string) *time.Location {
	if name == "" {
		return time.Local
	}
	if loc, err := time.LoadLocation(name); err == nil && loc != nil {
		return loc
	}
	// Fallback for Asia/Jakarta if tzdata/zoneinfo is missing
	if strings.EqualFold(name, "Asia/Jakarta") {
		return time.FixedZone("WIB", 7*3600) // UTC+7, no DST
	}
	// Last resort: local time (non-nil)
	return time.Local
}

func safeDate(year, month, day int, loc *time.Location) (time.Time, error) {
	d := time.Date(year, time.Month(month), day, 0, 0, 0, 0, loc)
	if d.Month() != time.Month(month) || d.Day() != day {
		return time.Time{}, fmt.Errorf("tanggal tidak valid")
	}
	return d, nil
}

func allSundays(year, month int, loc *time.Location) []time.Time {
	var res []time.Time
	for d := time.Date(year, time.Month(month), 1, 0, 0, 0, 0, loc); d.Month() == time.Month(month); d = d.AddDate(0, 0, 1) {
		if d.Weekday() == time.Sunday {
			res = append(res, d)
		}
	}
	return res
}

func sameDay(a, b time.Time) bool {
	ay, am, ad := a.Date()
	by, bm, bd := b.Date()
	return ay == by && am == bm && ad == bd
}

func cell(col, row int) string { ref, _ := excelize.CoordinatesToCellName(col, row); return ref }

func filterCandidates(people []Person, src string, mustPenatua bool) []string {
	key := normKey(src)
	m := map[string]struct{}{}
	for _, p := range people {
		if mustPenatua && !p.IsPenatua {
			continue
		}
		if mark, ok := p.Marks[key]; ok && mark {
			m[p.Name] = struct{}{}
		}
	}
	var res []string
	for n := range m {
		res = append(res, n)
	}
	sort.Strings(res)
	return res
}

func uniq(in []string) []string {
	m := map[string]struct{}{}
	var res []string
	for _, s := range in {
		if _, ok := m[s]; ok {
			continue
		}
		m[s] = struct{}{}
		res = append(res, s)
	}
	sort.Strings(res)
	return res
}

// helper to quiet unused var warnings in format string above
func jemaatNames(in []string) []string { return in }

func parseMonth(s string) (int, error) {
	m := map[string]int{"januari": 1, "februari": 2, "maret": 3, "april": 4, "mei": 5, "juni": 6, "juli": 7, "agustus": 8, "september": 9, "oktober": 10, "november": 11, "desember": 12}
	if n, ok := m[strings.ToLower(strings.TrimSpace(s))]; ok {
		return n, nil
	}
	var x int
	if _, err := fmt.Sscanf(s, "%d", &x); err == nil && x >= 1 && x <= 12 {
		return x, nil
	}
	return 0, fmt.Errorf("bulan tidak valid: %s", s)
}
func monthNameID(m int) string {
	names := []string{"", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"}
	if m >= 1 && m <= 12 {
		return names[m]
	}
	return "?"
}

// New: day name (ID)
func dayNameID(wd time.Weekday) string {
	switch wd {
	case time.Monday:
		return "Senin"
	case time.Tuesday:
		return "Selasa"
	case time.Wednesday:
		return "Rabu"
	case time.Thursday:
		return "Kamis"
	case time.Friday:
		return "Jumat"
	case time.Saturday:
		return "Sabtu"
	default:
		return "Minggu"
	}
}

// New: placeholder replacer
func replacePlaceholders(s string, d time.Time, loc *time.Location) string {
	day := dayNameID(d.Weekday())
	dd := fmt.Sprintf("%02d", d.Day())
	mon := monthNameID(int(d.Month()))
	yyyy := fmt.Sprintf("%04d", d.Year())
	out := s
	out = strings.ReplaceAll(out, "{Day}", day)
	out = strings.ReplaceAll(out, "{dd}", dd)
	// treat {MMM} and {MMMM} as full month name in ID
	out = strings.ReplaceAll(out, "{MMM}", mon)
	out = strings.ReplaceAll(out, "{MMMM}", mon)
	out = strings.ReplaceAll(out, "{yyyy}", yyyy)
	return out
}

// ==================== Pattern & Role Helpers ====================

func parsePattern(code string) (penatua, jemaat, total int, err error) {
	c := strings.ToLower(strings.TrimSpace(code))
	if len(c) < 2 {
		return 0, 0, 0, fmt.Errorf("kode '%s' tidak valid", code)
	}
	var n int
	var suf string
	if _, e := fmt.Sscanf(c, "%d%s", &n, &suf); e != nil {
		return 0, 0, 0, fmt.Errorf("kode '%s' tidak valid", code)
	}
	if n < 1 || n > 4 {
		return 0, 0, 0, fmt.Errorf("jumlah '%d' di luar batas 1..4", n)
	}
	switch n {
	case 1:
		if suf == "a" {
			return 1, 0, 1, nil
		}
		if suf == "b" {
			return 0, 1, 1, nil
		}
	case 2:
		switch suf {
		case "a":
			return 1, 1, 2, nil
		case "b":
			return 2, 0, 2, nil
		case "c":
			return 0, 2, 2, nil
		}
	case 3:
		switch suf {
		case "a":
			return 1, 2, 3, nil
		case "b":
			return 2, 1, 3, nil
		case "c":
			return 3, 0, 3, nil
		case "d":
			return 0, 3, 3, nil
		}
	case 4:
		switch suf {
		case "a":
			return 1, 3, 4, nil
		case "b":
			return 2, 2, 4, nil
		case "c":
			return 3, 1, 4, nil
		case "d":
			return 4, 0, 4, nil
		case "e":
			return 0, 4, 4, nil
		}
	}
	return 0, 0, 0, fmt.Errorf("kode '%s' tidak dikenali", code)
}

func isMajelisPendamping(role string) bool {
	r := strings.ToLower(role)
	return strings.Contains(r, "majel") && strings.Contains(r, "pend")
}

func defaultSlotsForRole(role, svc string, maxLektor, maxPro, maxMus int) int {
	low := strings.ToLower(strings.TrimSpace(role))
	if strings.Contains(low, "lektor") {
		return maxLektor
	}
	if strings.Contains(low, "prokantor") {
		return maxPro
	}
	if strings.Contains(low, "pemusik") {
		return maxMus
	}
	return 1
}
