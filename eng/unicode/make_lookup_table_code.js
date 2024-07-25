var fs = require('fs');
var s = fs.readFileSync("./temp/re_canon_tab.json", { encoding: "utf-8" });
var a = JSON.parse(s);

var entries = [];

for (let i = 0; i < a.length; ++i) {
    if (a[i] !== i) {
        entries.push({ s: i, e: i, delta: a[i] - i, step: 0});
    }
}

var entriesCompressed = [];
entriesCompressed[0] = { ...entries[0] };

for (let i = 1; i < entries.length; ++i) {
    const current = entries[i];
    const last = entriesCompressed[entriesCompressed.length - 1];
    if (current.delta === last.delta && current.s === last.e + 1 && (last.step === 0 || last.step === 1)) {
        last.e += 1; last.step = 1;
    } else if (current.delta === last.delta && current.s === last.e + 2 && (last.step === 0 || last.step === 2)) {
        last.e += 2; last.step = 2;
    } else {
        entriesCompressed.push({ ...current });
    }
}

const entries32 = [];
for (let i = 0; i < entriesCompressed.length; ++i) {
    const current = entriesCompressed[i];
    const delta = current.delta;
    const deltaunsigned = (delta >= 0 ? delta : delta + 0x10000) & 0xFFFF;
    const deltasigned = (deltaunsigned > 0x7FFF) ? deltaunsigned - 0x10000 : deltaunsigned;

    const step = current.step === 0 ? 1 : current.step;
    if ( (current.e - current.s) / step < 4) {
        for (j = current.s; j <= current.e; j += step) {
            entries32.push({ s: j, e: j, deltasigned, step });
        }
    } else {
        entries32.push({ s: current.s, e: current.e, deltasigned, step });
    }
}

function verify() {
    let aa = Array(entries.length);
    for (let i = 0; i <= 0xFFFF; ++i) aa[i] = i;
    for (let i = 0; i < entries32.length; ++i) {
        const current = entries32[i];
        const step = current.step === 0? 1 : current.step;
        for (let j = current.s; j <= current.e; j += step) {
            aa[j] = (aa[j] + current.deltasigned + 0x10000) & 0xFFFF;
        }
    }

    for (let i = 0; i <= 0xFFFF; ++i) if (aa[i] !== a[i]) {
        console.log(i, a[i], aa[i]);
        throw 0;
    }
}

verify();

var out = `Dim UnicodeCanonLookupTable(0 To ${a.length-1}) As Integer
Private Sub InititalizeUnicodeCanonLookupTable(ByRef t() As Integer)\n`;
for (let i = 0; i < entries32.length; ++i) {
    const current = entries32[i];
    const step = current.step === 0 ? 1 : current.step;
    if (current.s === current.e) {
        out += `\tt(${current.s}) = ${current.deltasigned}\n`;
    } else {
        const stepstr = current.step !== 1 ? " STEP " + current.step : "";
        out += `\tFOR i = ${current.s} TO ${current.e}${stepstr} : t(i) = ${current.deltasigned} : NEXT i\n`
    }
}
out += `End Sub\n`


var lst = -1;
var runs = [];
for (let i = 0; i < entriesCompressed.length; ++i) {
    if (entriesCompressed[i].s !== lst) runs.push(entriesCompressed[i].s);
    lst = entriesCompressed[i].e + 1
    runs.push(lst);
}

out += "\n";

out += `Dim UnicodeCanonRunsTable(0 To ${runs.length - 1}) As Long\n\n`

out += "Private Sub InitializeUnicodeCanonRunsTable(ByRef t() As Long)\n";

for (let i = 0; i < runs.length; ++i) {
    out += `\tt(${i}) = ${runs[i]}\n`;
}

out += "End Sub\n";
fs.writeFileSync("./temp/code.txt", out, { encoding: "utf-8" })
