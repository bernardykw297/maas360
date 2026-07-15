// State ID to abbreviation mapping
const stateIdToAbbr = {
    '01': 'AL', '02': 'AK', '04': 'AZ', '05': 'AR', '06': 'CA', '08': 'CO', '09': 'CT', '10': 'DE',
    '11': 'DC', '12': 'FL', '13': 'GA', '15': 'HI', '16': 'ID', '17': 'IL', '18': 'IN', '19': 'IA',
    '20': 'KS', '21': 'KY', '22': 'LA', '23': 'ME', '24': 'MD', '25': 'MA', '26': 'MI', '27': 'MN',
    '28': 'MS', '29': 'MO', '30': 'MT', '31': 'NE', '32': 'NV', '33': 'NH', '34': 'NJ', '35': 'NM',
    '36': 'NY', '37': 'NC', '38': 'ND', '39': 'OH', '40': 'OK', '41': 'OR', '42': 'PA', '44': 'RI',
    '45': 'SC', '46': 'SD', '47': 'TN', '48': 'TX', '49': 'UT', '50': 'VT', '51': 'VA', '53': 'WA',
    '54': 'WV', '55': 'WI', '56': 'WY', '72': 'PR'
};

let selectedState = null;

// Initialize the map
async function initMap() {
    const svg = d3.select('#us-map');
    const width = 960;
    const height = 600;

    const projection = d3.geoAlbersUsa().scale(1200).translate([width / 2, height / 2]);
    const path = d3.geoPath().projection(projection);

    try {
        const us = await d3.json('https://cdn.jsdelivr.net/npm/us-atlas@3/states-10m.json');
        const states = topojson.feature(us, us.objects.states);

        svg.selectAll('path')
            .data(states.features)
            .enter()
            .append('path')
            .attr('d', path)
            .attr('class', d => {
                const stateId = d.id.toString().padStart(2, '0');
                const stateCode = stateIdToAbbr[stateId];
                return stateRepData[stateCode] ? 'state has-data' : 'state';
            })
            .attr('data-state', d => {
                const stateId = d.id.toString().padStart(2, '0');
                return stateIdToAbbr[stateId];
            })
            .on('click', function(event, d) {
                const stateId = d.id.toString().padStart(2, '0');
                const stateCode = stateIdToAbbr[stateId];
                handleStateClick(stateCode);
            });
    } catch (error) {
        console.error('Error loading map:', error);
        document.getElementById('state-info').innerHTML = '<div class="no-data"><p><strong>Error Loading Map</strong></p><p>Unable to load US map data.</p></div>';
    }
}

function handleStateClick(stateCode) {
    if (!stateCode) return;
    d3.selectAll('.state').classed('selected', false);
    selectedState = stateCode;
    d3.selectAll(`[data-state="${stateCode}"]`).classed('selected', true);
    displayStateInfo(stateCode);
}

function displayStateInfo(stateCode) {
    const infoPanel = document.getElementById('state-info');
    const data = stateRepData[stateCode];
    
    if (!data) {
        infoPanel.innerHTML = '<div class="no-data"><p><strong>' + getStateName(stateCode) + '</strong></p><p>No representative data available.</p></div>';
        return;
    }
    
    let html = '<h2 class="state-name">' + data.name + '</h2>';
    
    html += '<div class="rep-section"><h3>Sales Representative</h3><div class="rep-info">';
    html += '<div class="rep-detail"><strong>Name</strong><span>' + data.rep + '</span></div>';
    if (data.repEmail) html += '<div class="rep-detail"><strong>Email</strong><span><a href="mailto:' + data.repEmail + '">' + data.repEmail + '</a></span></div>';
    if (data.repPhone) html += '<div class="rep-detail"><strong>Phone</strong><span>' + data.repPhone + '</span></div>';
    html += '</div></div>';
    
    html += '<div class="rep-section"><h3>Telco Specialist</h3>';

    // Verizon
    html += '<div class="telco-sub-section"><h4 class="telco-carrier-label">Verizon</h4><div class="rep-info">';
    html += '<div class="rep-detail"><strong>Name</strong><span>' + (data.telcoSpecialist || 'Open') + '</span></div>';
    if (data.telcoEmail) html += '<div class="rep-detail"><strong>Email</strong><span><a href="mailto:' + data.telcoEmail + '">' + data.telcoEmail + '</a></span></div>';
    if (data.telcoPhone) html += '<div class="rep-detail"><strong>Phone</strong><span>' + data.telcoPhone + '</span></div>';
    html += '</div></div>';

    // AT&T
    html += '<div class="telco-sub-section"><h4 class="telco-carrier-label">AT&amp;T</h4><div class="rep-info">';
    html += '<div class="rep-detail"><strong>Name</strong><span>' + (data.attSpecialist || 'Open') + '</span></div>';
    if (data.attEmail) html += '<div class="rep-detail"><strong>Email</strong><span><a href="mailto:' + data.attEmail + '">' + data.attEmail + '</a></span></div>';
    if (data.attPhone) html += '<div class="rep-detail"><strong>Phone</strong><span>' + data.attPhone + '</span></div>';
    html += '</div></div>';

    // T-Mobile
    html += '<div class="telco-sub-section"><h4 class="telco-carrier-label">T-Mobile</h4><div class="rep-info">';
    html += '<div class="rep-detail"><strong>Name</strong><span>' + (data.tmobileSpecialist || 'Open') + '</span></div>';
    if (data.tmobileEmail) html += '<div class="rep-detail"><strong>Email</strong><span><a href="mailto:' + data.tmobileEmail + '">' + data.tmobileEmail + '</a></span></div>';
    if (data.tmobilePhone) html += '<div class="rep-detail"><strong>Phone</strong><span>' + data.tmobilePhone + '</span></div>';
    html += '</div></div>';

    html += '</div>';
    
    html += '<div class="rep-section"><h3>Technical Representative</h3><div class="rep-info">';
    html += '<div class="rep-detail"><strong>Name</strong><span>' + data.techRep + '</span></div>';
    if (data.techRepEmail) html += '<div class="rep-detail"><strong>Email</strong><span><a href="mailto:' + data.techRepEmail + '">' + data.techRepEmail + '</a></span></div>';
    if (data.techRepPhone) html += '<div class="rep-detail"><strong>Phone</strong><span>' + data.techRepPhone + '</span></div>';
    html += '</div></div>';
    
    html += '<div class="rep-section"><h3>Manager</h3><div class="rep-info">';
    html += '<div class="rep-detail"><strong>Name</strong><span>' + data.manager + (data.managerTitle ? ' - ' + data.managerTitle : '') + '</span></div>';
    if (data.managerEmail) html += '<div class="rep-detail"><strong>Email</strong><span><a href="mailto:' + data.managerEmail + '">' + data.managerEmail + '</a></span></div>';
    if (data.managerPhone) html += '<div class="rep-detail"><strong>Phone</strong><span>' + data.managerPhone + '</span></div>';
    html += '</div></div>';
    
    html += '<div class="rep-section"><h3>Technical Manager</h3><div class="rep-info">';
    html += '<div class="rep-detail"><strong>Name</strong><span>' + data.techManager + '</span></div>';
    if (data.techManagerEmail) html += '<div class="rep-detail"><strong>Email</strong><span><a href="mailto:' + data.techManagerEmail + '">' + data.techManagerEmail + '</a></span></div>';
    if (data.techManagerPhone) html += '<div class="rep-detail"><strong>Phone</strong><span>' + data.techManagerPhone + '</span></div>';
    html += '</div></div>';
    
    infoPanel.innerHTML = html;
}

function getStateName(code) {
    const names = {'AL':'Alabama','AK':'Alaska','AZ':'Arizona','AR':'Arkansas','CA':'California','CO':'Colorado','CT':'Connecticut','DE':'Delaware','FL':'Florida','GA':'Georgia','HI':'Hawaii','ID':'Idaho','IL':'Illinois','IN':'Indiana','IA':'Iowa','KS':'Kansas','KY':'Kentucky','LA':'Louisiana','ME':'Maine','MD':'Maryland','MA':'Massachusetts','MI':'Michigan','MN':'Minnesota','MS':'Mississippi','MO':'Missouri','MT':'Montana','NE':'Nebraska','NV':'Nevada','NH':'New Hampshire','NJ':'New Jersey','NM':'New Mexico','NY':'New York','NC':'North Carolina','ND':'North Dakota','OH':'Ohio','OK':'Oklahoma','OR':'Oregon','PA':'Pennsylvania','RI':'Rhode Island','SC':'South Carolina','SD':'South Dakota','TN':'Tennessee','TX':'Texas','UT':'Utah','VT':'Vermont','VA':'Virginia','WA':'Washington','WV':'West Virginia','WI':'Wisconsin','WY':'Wyoming','DC':'District of Columbia'};
    return names[code] || code;
}

document.addEventListener('DOMContentLoaded', initMap);
