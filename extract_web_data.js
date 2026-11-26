const fs = require('fs');
const html = fs.readFileSync('index.html', 'utf8');

const startIdx = html.indexOf('const visualConfigs = {');
const endIdx = html.indexOf('function openStationWorksheetModal');

const scriptPart = html.slice(startIdx, endIdx);

const jsToRun = `
${scriptPart}

const results = {};
for (let stId = 1; stId <= 4; stId++) {
    const data = visualConfigs[stId];
    const optSols = data.solutions.filter(s => s.isOptimal);
    const comments = observationComments[stId] || [];
    
    const steps = [
        { name: "Ban đầu (Hiện trạng)", acts: [false, false, false, false], solName: "Chưa kích hoạt giải pháp nào" },
        { name: "Bước 1", acts: [true, false, false, false], solName: optSols[0] ? optSols[0].title : "" },
        { name: "Bước 2", acts: [true, true, false, false], solName: optSols[1] ? optSols[1].title : "" },
        { name: "Bước 3", acts: [true, true, true, false], solName: optSols[2] ? optSols[2].title : "" },
        { name: "Bước 4", acts: [true, true, true, true], solName: optSols[3] ? optSols[3].title : "" }
    ];

    const stepRows = steps.map((st, idx) => {
        const tele = computeStationTelemetry(stId, st.acts);
        return {
            step: st.name,
            solution: st.solName,
            co2: tele.co2,
            light: tele.light,
            temp: tele.temp,
            status: tele.status,
            comment: comments[idx] || ""
        };
    });

    const initTele = computeStationTelemetry(stId, [false, false, false, false]);
    const finalTele = computeStationTelemetry(stId, [true, true, true, true]);

    results[stId] = {
        title: data.title,
        subtitle: data.subtitle,
        badge: data.badge,
        challenge: data.challenge,
        steps: stepRows,
        initTele: initTele,
        finalTele: finalTele
    };
}

fs.writeFileSync('extracted_station_data.json', JSON.stringify(results, null, 2), 'utf8');
console.log('Successfully extracted 100% exact web data to extracted_station_data.json!');
`;

const vm = require('vm');
const sandbox = {
    fs: fs,
    console: console,
    currentStationId: 1,
    stationState: { 1: { chosenSolutions: [] }, 2: { chosenSolutions: [] }, 3: { chosenSolutions: [] }, 4: { chosenSolutions: [] } },
    document: {
        getElementById: () => ({ innerText: '', value: '', classList: { add: () => {}, remove: () => {} }, style: {} })
    },
    window: { addEventListener: () => {} },
    localStorage: { getItem: () => null, setItem: () => {} },
    alert: () => {},
    setTimeout: (fn) => fn(),
    clearTimeout: () => {}
};

vm.createContext(sandbox);
vm.runInContext(jsToRun, sandbox);
