<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<title>市價徵收查估作業-影響地價區域因素明細表(住宅用地)</title>

<link rel="stylesheet" href="https://unpkg.com/leaflet/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet/dist/leaflet.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script>
<!-- 幾何計算核心 -->
<script src="https://cdn.jsdelivr.net/npm/@turf/turf@6/turf.min.js"></script>
<!-- 前端 SHP 轉檔打包工具 -->
<script src="https://unpkg.com/shp-write@latest/shpwrite.js"></script>
<!-- 新增 DEM GeoTIFF -->
<script src="https://unpkg.com/geotiff/dist-browser/geotiff.js"></script>    
<!-- 關鍵擴充：引入 JSZip 與 shpjs，讓前端具備真實解碼 SHP 壓縮檔的能力 -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.7.1/jszip.min.js"></script>
<script src="https://unpkg.com/shpjs@latest/dist/shp.js"></script>

<style>
body { margin:0; font-family: Arial, "Microsoft JhengHei"; }
#map { height:100vh; }

/* UI 面板 */
#panel{
position:absolute;
top:12px;
right:12px;
width:320px;
background: linear-gradient(135deg,#fffdf5,#fef3c7);
padding:14px;
z-index:1000;
border-radius:16px;
box-shadow:0 10px 25px rgba(0,0,0,0.25);
font-size:13px;
border:1px solid #fcd34d;
}  

#panel h3{margin:0 0 8px 0;color:#1e3a8a;}

#panel input{
width:95%;
padding:6px;
margin:3px 0;
border-radius:8px;
border:1px solid #cbd5e1;
}

#panel label{
display:grid;
grid-template-columns: 18px 1fr;
align-items:center;
column-gap:8px;
padding:4px 6px;
border-radius:8px;
cursor:pointer;
}

#panel label:hover{background:#e0ecff;}

#panel button{
width:100%;
margin-top:6px;
padding:8px;
border:none;
border-radius:10px;
cursor:pointer;
font-weight:bold;
background:#2563eb;color:white;
}

#panel button:disabled {
background:#94a3b8;
cursor:not-allowed;
}

#panel button.clear-btn{
background:#10b981;
}

#resultPanel{
position:absolute;
top:120px;
left:120px;
width:40%;
max-width:600px;
background:white;
border-radius:14px;
box-shadow:0 10px 25px rgba(0,0,0,0.25);
z-index:999;
overflow:hidden;
}

#resultHeader{
background:#2563eb;
color:white;
padding:8px;
font-weight:bold;
cursor:grab;
user-select:none;
display:flex;
justify-content:space-between;
align-items:center;
}

#resultHeader button{
background:#10b981;
border:none;
color:white;
padding:4px 8px;
border-radius:6px;
cursor:pointer;
font-size:12px;
}

#resultTableWrapper{
max-height:180px;
overflow-y:auto;
}

table{
width:100%;
border-collapse: collapse;
font-size:12px;
table-layout: fixed;
}

td,th{
border:1px solid #e5e7eb;
padding:6px;
text-align:center;
vertical-align:middle;
}

th{
background:#dbeafe;
position:sticky;
top:0;
}

.highlight{
background:#fde68a !important;
}

#routeInfo{
margin-top:10px;
padding:8px;
background:#eff6ff;
border-radius:10px;
color:#1d4ed8;
font-weight:bold;
}

/* 繪製控制項工具列擴充 */
.draw-control-bar {
    background: white;
    padding: 6px;
    border-radius: 8px;
    box-shadow: 0 1px 5px rgba(0,0,0,0.4);
    display: flex;
    flex-direction: column;
    gap: 5px;
}
.draw-btn {
    background: #ffffff;
    color: #333;
    border: 1px solid #ccc;
    padding: 6px 10px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 12px;
    font-weight: bold;
    text-align: center;
    white-space: nowrap;
}
.draw-btn:hover { background: #f4f4f4; }
.draw-btn.active { background: #ef4444; color: white; border-color: #dc2626; }
.shp-btn { background: #4f46e5; color: white; border: none; }
.shp-btn:hover { background: #4338ca; }
.shp-import-label {
    background: #059669; color: white; border: none;
    padding: 6px 10px; border-radius: 4px; cursor: pointer;
    font-size: 12px; font-weight: bold; text-align: center;
}
.shp-import-label:hover { background: #047857; }

/* 拖曳編輯控制點樣式 */
.edit-handle {
    width: 12px;
    height: 12px;
    background: #ffffff;
    border: 2px solid #dc2626;
    border-radius: 50%;
    box-shadow: 0 1px 4px rgba(0,0,0,0.3);
}

/* 多邊形列表面板 */
#polygonListPanel {
    position:absolute;
    top:12px;
    left:12px;
    width:280px;
    background:#ffffff;
    border-radius:12px;
    box-shadow:0 10px 25px rgba(0,0,0,0.25);
    z-index:1200;
    overflow:hidden;
    font-size:12px;
}

#polygonListHeader {
    background:#7c3aed;
    color:white;
    padding:8px;
    font-weight:bold;
    cursor:grab;
    display:flex;
    justify-content:space-between;
    align-items:center;
    user-select:none;
}

#polygonListHeader button {
    background:#10b981;
    border:none;
    color:white;
    padding:4px 8px;
    border-radius:6px;
    cursor:pointer;
    font-size:11px;
}

#polygonListContent {
    padding:8px;
    background:#f3e8ff;
    max-height:300px;
    overflow-y:auto;
}

.polygon-item {
    padding:6px;
    margin:4px 0;
    background:white;
    border-radius:6px;
    border:1px solid #ddd;
    cursor:pointer;
    display:flex;
    justify-content:space-between;
    align-items:center;
}

.polygon-item:hover {
    background:#f0f0f0;
    border-color:#7c3aed;
}

.polygon-item.selected {
    background:#dbeafe;
    border-color:#2563eb;
    font-weight:bold;
}

.polygon-item-actions {
    display:flex;
    gap:4px;
}

.polygon-item-actions button {
    background:#ef4444;
    color:white;
    border:none;
    padding:3px 6px;
    border-radius:3px;
    cursor:pointer;
    font-size:10px;
}

.polygon-item-actions button:hover {
    background:#dc2626;
}

/* 量測信息面板 */
#measurementPanel {
    position:absolute;
    bottom:100px;
    left:12px;
    width:320px;
    background:white;
    border-radius:12px;
    box-shadow:0 10px 25px rgba(0,0,0,0.25);
    z-index:999;
    overflow:hidden;
}

#measurementHeader {
    background:#1e40af;
    color:white;
    padding:8px;
    font-weight:bold;
    cursor:grab;
    display:flex;
    justify-content:space-between;
    align-items:center;
}

#measurementContent {
    padding:8px;
    max-height:250px;
    overflow-y:auto;
    font-size:11px;
    line-height:1.6;
}

#slopePanel{
position:absolute;
top:12px;
left:12px;
width:260px;
background:#ffffff;
border-radius:12px;
box-shadow:0 10px 25px rgba(0,0,0,0.25);
z-index:1200;
overflow:hidden;
font-size:12px;
}

#slopeHeader{
background:#7c3aed;
color:white;
padding:8px;
font-weight:bold;
cursor:grab;
display:flex;
justify-content:space-between;
align-items:center;
user-select:none;
}

#slopeHeader button{
background:#10b981;
border:none;
color:white;
padding:4px 8px;
border-radius:6px;
cursor:pointer;
font-size:11px;
}

#slopeContent{
padding:8px;
background:#f3e8ff;
color:#5b21b6;
line-height:1.6;
}

</style>
</head>
<body>

<div style="position:absolute;bottom:8px;left:12px;font-size:11px;color:#666;background:rgba(255,255,255,0.8);padding:4px 8px;border-radius:6px;z-index:1000;">
圖資來源 © OpenStreetMap contributors ／ 國土測繪中心(NLSC)
路網分析：OSRM Demo Server ／ 空間計算：Turf.js & shp-write
</div>

<div id="panel">
<h3>影響地價區域因素明細表(住宅用地)</h3>

緯度:<input id="lat" value="25.012098"><br>
經度:<input id="lon" value="121.465675"><br>
半徑(m):<input id="radius" value="1000"><br>

<h3>類型</h3>

<label><input type="checkbox" value="station">交通運輸-接近大型車站之程度</label>
<label><input type="checkbox" value="bus_stop">交通運輸-站牌之接近程度或密集程度</label>
<label><input type="checkbox" value="motorway_junction">交通運輸-交流道之有無及接近交流道之程度</label>

<label><input type="checkbox" value="school" checked>公共建設-接近學校之程度(門牌大門優先)</label>
<label><input type="checkbox" value="marketplace">公共建設-接近市場之程度(傳統市場、超級市場、超大型購物中心)</label>
<label><input type="checkbox" value="park">公共建設-接近公園(里鄰公園、一般公園)、廣場、徒步區之程度</label>
<label><input type="checkbox" value="tourism">公共建設-接近觀光遊憩設施之程度</label>
<label><input type="checkbox" value="parking">公共建設-停車場地之便利程度</label>
<label><input type="checkbox" value="bank">公共建設-接近服務性設施的程度(郵局、銀行、醫院、機關等設施)</label>
<label><input type="checkbox" value="gas">特殊設施-電業設施及公用氣體燃料設施之有無極接近程度</label>
<label><input type="checkbox" value="cemetery">特殊建設-殯葬設施之有無及接近程度</label>  
<label><input type="checkbox" value="waste_disposal">特殊建設-廢棄物處理設施之有無及接近程度</label>
<label><input type="checkbox" value="pollution">環境汙染-水汙染、噪音汙染、廢氣汙染、廢棄物汙染等之有無及接近程度</label>
<button id="query-btn" onclick="runQuery()">查詢</button>
<div style="display:flex;gap:6px;margin-top:6px;">

<button type="button"
onclick="getCurrentLocation()"
style="
flex:1;
background:#0ea5e9;
color:white;
border:none;
border-radius:10px;
font-weight:bold;
cursor:pointer;
">
取得現在位置
</button>
</div>

<button class="clear-btn" onclick="clearRoute()">清除路線距離</button>

<div id="routeInfo">路線距離：尚未計算</div>
<hr style="margin:10px 0;">

   
</div>

<div id="resultPanel">
<div id="resultHeader">
查詢結果
<button onclick="exportExcel()">Excel</button>
</div>

<div id="resultTableWrapper">
<table>
<thead>
<tr>
<th>名稱</th>
<th>類型</th>
<th onclick="sortResults('distance')" style="cursor:pointer;" title="點擊排序">直線距離 ↕</th>
<th onclick="sortResults('route')" style="cursor:pointer;" title="點擊排序">路線距離 ↕</th>
</tr>
</thead>
<tbody id="tbody"></tbody>
</table>
</div>
</div>

<!-- 多邊形列表面板 -->
<div id="polygonListPanel">
    <div id="polygonListHeader">
        🗺️ 多邊形管理
        <button onclick="addNewPolygon()">➕ 新增</button>
    </div>
    <div id="polygonListContent">
        <div id="polygonListItems"></div>
    </div>
</div>

<!-- 量測信息面板 -->
<div id="measurementPanel" style="display:none;">
    <div id="measurementHeader">
        📏 形心距離量測
        <span style="cursor:pointer;" onclick="toggleMeasurementPanel()">➖</span>
    </div>
    <div id="measurementContent">
        <div id="measurementInfo">請選擇多邊形進行量測</div>
    </div>
</div>

<div id="slopePanel">
    <div id="slopeHeader">
        🌄 地形坡度分析（可拖曳）
        <button onclick="analyzeSlope()">執行</button>
    </div>

    <div id="slopeContent">
        <div id="slopeInfo">尚未分析坡度</div>
    </div>
</div>

<div id="map"></div>

<script src="fixes/fix-fetch-and-measurements.js"></script>

<script>
let map = L.map('map').setView([25.012098, 121.465675], 15);

/* =========================================================
   GIS 多底圖架構（合法替代 Google Maps）
   ========================================================= */

/* =========================================================
   1. NLSC 電子地圖（預設）
   ========================================================= */
const nlscEMAP = L.tileLayer(
'https://wmts.nlsc.gov.tw/wmts/EMAP/default/EPSG:3857/{z}/{y}/{x}',
{
    maxZoom:19,
    attribution:'© NLSC'
});

/* =========================================================
   2. NLSC 正射影像
   ========================================================= */
const nlscPHOTO = L.tileLayer(
'https://wmts.nlsc.gov.tw/wmts/PHOTO2/default/EPSG:3857/{z}/{y}/{x}',
{
    maxZoom:20,
    attribution:'© NLSC 正射影像'
});

/* =========================================================
   3. OpenStreetMap
   ========================================================= */
const osmLayer = L.tileLayer(
'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
{
    maxZoom:19,
    attribution:'© OpenStreetMap contributors'
});

/* =========================================================
   4. ESRI World Imagery（合法衛星圖）
   ========================================================= */
const esriWorldImagery = L.tileLayer(
'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
{
    maxZoom:19,
    attribution:'© Esri World Imagery'
});

/* =========================================================
   預設底圖
   ========================================================= */
nlscEMAP.addTo(map);

/* =========================================================
   地籍圖（獨立 overlay）
   ========================================================= */
const cadastralLayer = L.tileLayer(
'https://wmts.nlsc.gov.tw/wmts/CADASTRAL/default/GoogleMapsCompatible/{z}/{y}/{x}',
{
    maxZoom:20,
    opacity:0.75,
    attribution:'© NLSC 地籍圖'
}
);

cadastralLayer.addTo(map);

/* =========================================================
   底圖切換
   ========================================================= */
const baseMaps = {
    "NLSC 電子地圖": nlscEMAP,
    "NLSC 正射影像": nlscPHOTO,
    "OpenStreetMap": osmLayer,
    "ESRI 衛星影像": esriWorldImagery
};

/* =========================================================
   GIS Overlay
   ========================================================= */
/* =========================================================
   NLSC DEM 地形暈渲圖（Hillshade）
   ========================================================= */

const nlscDEM = L.tileLayer(
'https://wmts.nlsc.gov.tw/wmts/EMAP5_OPENDATA/default/EPSG:3857/{z}/{y}/{x}',
{
    maxZoom:18,
    opacity:0.45,
    attribution:'© NLSC DEM'
});

/* =========================================================
   GIS Overlay
   ========================================================= */

const overlayMaps = {
    "地籍圖": cadastralLayer,
    "DEM 地形": nlscDEM
};

/* =========================================================
   圖層控制器
   ========================================================= */
L.control.layers(baseMaps, overlayMaps,{
    collapsed:false,
    position:'topright'
}).addTo(map);

/* 比例尺 */
const scaleControl = L.control({ position: 'bottomleft' });

scaleControl.onAdd = function () {
this._div = L.DomUtil.create('div', 'scale-denominator');
this._div.style.background = 'rgba(255,255,255,0.9)';
this._div.style.padding = '4px 8px';
this._div.style.borderRadius = '6px';
this._div.style.boxShadow = '0 1px 5px rgba(0,0,0,0.3)';
this._div.style.fontSize = '12px';
this._div.style.fontWeight = 'bold';
this._div.style.marginBottom = '42px';
this.update();
return this._div;
};

scaleControl.update = function () {
const center = map.getCenter();
const zoom = map.getZoom();
const metersPerPixel = 40075016.686 * Math.abs(Math.cos(center.lat * Math.PI / 180)) / Math.pow(2, zoom + 8);
const dpi = 96;
const scale = metersPerPixel * dpi / 0.0254;
const rounded = Math.round(scale / 100) * 100;
this._div.innerHTML = `比例尺 1:${rounded.toLocaleString()}`;
};

scaleControl.addTo(map);
map.on('zoomend moveend', () => {
    scaleControl.update();
});

let markerLayer=L.layerGroup().addTo(map);
  
const latInput=document.getElementById("lat");
const lonInput=document.getElementById("lon");
const radiusInput=document.getElementById("radius");
const tbody=document.getElementById("tbody");
const routeInfo=document.getElementById("routeInfo");
const queryBtn=document.getElementById("query-btn");

let userMarker=null;
let origin=null;
let results=[];

let sortState = { distance:true, route:true };
let routeLayer=null;

origin=[25.012098,121.465675];

userMarker=L.marker(origin,{
icon:L.divIcon({
className:'',
html: `<div style="width:18px;height:18px;background:#ef4444;border:3px solid #ffffff;border-radius:50%;box-shadow:0 0 0 2px rgba(239,68,68,0.35);"></div>`,
iconSize:[18,18],
iconAnchor:[9,9]
}),
interactive:false
}).addTo(map);


/* ====================================================================
    ★ 改進版：多邊形繪製與管理核心
    ==================================================================== */

// 多邊形集合
let polygons = [];
let selectedPolygonId = null;
let isDrawingPolygon = false;
let currentDrawPoints = [];       
let previewPolyline = null;
let drawMarkers = L.layerGroup().addTo(map);

// 多邊形顏色池
const polygonColors = ['#dc2626', '#2563eb', '#059669', '#ea580c', '#a855f7', '#0891b2', '#7c2d12', '#831843'];
let colorIndex = 0;

// 多邊形物件結構
class Polygon {
    constructor(id, points, color) {
        this.id = id;
        this.points = points; // [[lat, lon], ...]
        this.color = color;
        this.leafletPolygon = null;
        this.editMarkers = L.layerGroup();
        this.centroid = null;
        this.area = 0;
        this.isSelected = false;
        this.measurementData = null;
        this.calculateCentroid();
    }

    calculateCentroid() {
        if (this.points.length < 3) return;
        
        let coordinates = this.points.map(p => [p[1], p[0]]);
        if (coordinates[0][0] !== coordinates[coordinates.length-1][0] || 
            coordinates[0][1] !== coordinates[coordinates.length-1][1]) {
            coordinates.push([this.points[0][1], this.points[0][0]]);
        }
        
        let turfPoly = turf.polygon([coordinates]);
        let centroidObj = turf.centroid(turfPoly);
        
        this.centroid = [centroidObj.geometry.coordinates[1], centroidObj.geometry.coordinates[0]];
        this.area = turf.area(turfPoly);
    }

    render() {
        if (this.leafletPolygon) map.removeLayer(this.leafletPolygon);
        this.editMarkers.clearLayers();

        this.leafletPolygon = L.polygon(this.points, {
            color: this.color,
            weight: this.isSelected ? 4 : 3,
            fillColor: this.color,
            fillOpacity: 0.3,
            className: this.isSelected ? 'selected-polygon' : ''
        }).addTo(map);

        if (this.isSelected) {
            this.renderEditHandles();
        }

        this.leafletPolygon.bindPopup(`
            <div style="font-size:12px; line-height:1.4;">
                <b style="color:${this.color};">多邊形 #${this.id}</b><br>
                面積: ${(this.area/10000).toFixed(2)} 公頃<br>
                頂點數: ${this.points.length}<br>
                形心: ${this.centroid[0].toFixed(6)}, ${this.centroid[1].toFixed(6)}
            </div>
        `);
    }

    renderEditHandles() {
        this.points.forEach((pt, index) => {
            let marker = L.marker(pt, {
                icon: L.divIcon({
                    className: '',
                    html: '<div class="edit-handle"></div>',
                    iconSize: [12, 12],
                    iconAnchor: [6, 6]
                }),
                draggable: true
            });
            
            marker.on('drag', (e) => {
                let newLatLng = e.target.getLatLng();
                this.points[index] = [newLatLng.lat, newLatLng.lng];
                this.calculateCentroid();
                this.render();
            });
            
            marker.on('dragend', () => {
                this.calculateCentroid();
                updateMeasurementForPolygon(this);
            });

            marker.addTo(this.editMarkers);
        });
        
        this.editMarkers.addTo(map);
    }

    delete() {
        if (this.leafletPolygon) map.removeLayer(this.leafletPolygon);
        this.editMarkers.clearLayers();
        polygons = polygons.filter(p => p.id !== this.id);
        if (selectedPolygonId === this.id) {
            selectedPolygonId = null;
        }
        updatePolygonList();
    }

    select() {
        if (selectedPolygonId !== null && selectedPolygonId !== this.id) {
            let oldPoly = polygons.find(p => p.id === selectedPolygonId);
            if (oldPoly) {
                oldPoly.isSelected = false;
                oldPoly.render();
            }
        }
        this.isSelected = true;
        selectedPolygonId = this.id;
        this.render();
        updateMeasurementForPolygon(this);
        map.fitBounds(L.polygon(this.points).getBounds());
    }

    toGeoJSON() {
        let ring = this.points.map(p => [p[1], p[0]]);
        ring.push([this.points[0][1], this.points[0][0]]);
        return {
            type: "Feature",
            geometry: { type: "Polygon", coordinates: [ring] },
            properties: { 
                id: this.id,
                name: `Polygon_${this.id}`,
                area: this.area,
                centroid: this.centroid
            }
        };
    }

    static fromGeoJSON(feature, id) {
        let ring = feature.geometry.coordinates[0];
        if (ring.length > 1 && ring[0][0] === ring[ring.length-1][0] && ring[0][1] === ring[ring.length-1][1]) {
            ring.pop();
        }
        let points = ring.map(coord => [coord[1], coord[0]]);
        let color = polygonColors[colorIndex % polygonColors.length];
        colorIndex++;
        return new Polygon(id, points, color);
    }
}

// 繪製控制條
const drawControl = L.control({ position: 'topleft' });
drawControl.onAdd = function() {
    const div = L.DomUtil.create('div', 'draw-control-bar');
    div.innerHTML = `
        <button id="btn-start-draw" class="draw-btn" onclick="toggleDrawMode(this)">📐 新增多邊形</button>
        <button id="btn-end-draw" class="draw-btn" onclick="finishPolygon()" style="display:none; background:#10b981; color:white;">✓ 完成閉合</button>
        <button id="btn-export-shp" class="draw-btn shp-btn" onclick="exportAllPolygonsSHP()" style="display:none;">📥 匯出全部 SHP</button>
        <label id="lbl-import-shp" class="shp-import-label">📤 匯入範圍 (支援 SHP.zip 或 GeoJSON)
            <input type="file" id="shp-file-input" accept=".zip,.geojson,.json" style="display:none;" onchange="importPolygonSHP(event)">
        </label>
        <button id="btn-clear-draw" class="draw-btn" onclick="cancelDrawing()" style="display:none; background:#6b7280; color:white;">✕ 取消繪製</button>
    `;
    L.DomEvent.disableClickPropagation(div);
    return div;
};
drawControl.addTo(map);

const resPanel = document.getElementById("resultPanel");
const sidePanel = document.getElementById("panel");
L.DomEvent.disableClickPropagation(resPanel);
L.DomEvent.disableClickPropagation(sidePanel);

function toggleDrawMode(btn) {
    isDrawingPolygon = !isDrawingPolygon;
    if (isDrawingPolygon) {
        btn.classList.add('active');
        btn.innerText = "🛑 停止點選";
        document.getElementById('btn-end-draw').style.display = 'block';
        document.getElementById('btn-clear-draw').style.display = 'block';
        currentDrawPoints = [];
        if (previewPolyline) { map.removeLayer(previewPolyline); previewPolyline = null; }
        drawMarkers.clearLayers();
    } else {
        btn.classList.remove('active');
        btn.innerText = "📐 新增多邊形";
        if (currentDrawPoints.length === 0) {
            document.getElementById('btn-end-draw').style.display = 'none';
            document.getElementById('btn-clear-draw').style.display = 'none';
        }
    }
}

function cancelDrawing() {
    isDrawingPolygon = false;
    currentDrawPoints = [];
    if (previewPolyline) { map.removeLayer(previewPolyline); previewPolyline = null; }
    drawMarkers.clearLayers();
    
    const btn = document.getElementById('btn-start-draw');
    btn.classList.remove('active');
    btn.innerText = "📐 新增多邊形";
    document.getElementById('btn-end-draw').style.display = 'none';
    document.getElementById('btn-clear-draw').style.display = 'none';
}

function finishPolygon() {
    if (currentDrawPoints.length < 3) {
        alert("請至少點選 3 個以上的頂點才能閉合範圍！");
        return;
    }
    if (previewPolyline) { map.removeLayer(previewPolyline); previewPolyline = null; }
    
    // 建立新多邊形
    let newId = polygons.length > 0 ? Math.max(...polygons.map(p => p.id)) + 1 : 1;
    let color = polygonColors[colorIndex % polygonColors.length];
    colorIndex++;
    
    let newPoly = new Polygon(newId, [...currentDrawPoints], color);
    polygons.push(newPoly);
    newPoly.select();
    newPoly.render();
    
    updatePolygonList();
    
    cancelDrawing();
    document.getElementById('btn-export-shp').style.display = 'block';
}

function addNewPolygon() {
    const btn = document.getElementById('btn-start-draw');
    if (!isDrawingPolygon) {
        toggleDrawMode(btn);
    }
}

function updatePolygonList() {
    const listDiv = document.getElementById('polygonListItems');
    listDiv.innerHTML = '';
    
    if (polygons.length === 0) {
        listDiv.innerHTML = '<div style="padding:8px; color:#999;">尚無多邊形</div>';
        return;
    }
    
    polygons.forEach(poly => {
        const item = document.createElement('div');
        item.className = `polygon-item ${poly.isSelected ? 'selected' : ''}`;
        item.innerHTML = `
            <div onclick="selectPolygon(${poly.id})" style="flex:1; cursor:pointer;">
                <div style="color:${poly.color}; font-weight:bold;">多邊形 #${poly.id}</div>
                <div style="font-size:10px; color:#666;">頂點: ${poly.points.length} | 面積: ${(poly.area/10000).toFixed(2)} 公頃</div>
            </div>
            <div class="polygon-item-actions">
                <button onclick="event.stopPropagation(); selectPolygon(${poly.id})">選取</button>
                <button onclick="event.stopPropagation(); deletePolygon(${poly.id})">刪除</button>
            </div>
        `;
        listDiv.appendChild(item);
    });
}

function selectPolygon(id) {
    let poly = polygons.find(p => p.id === id);
    if (poly) {
        poly.select();
        updatePolygonList();
    }
}

function deletePolygon(id) {
    if (confirm('確認刪除此多邊形？')) {
        let poly = polygons.find(p => p.id === id);
        if (poly) {
            poly.delete();
            updatePolygonList();
        }
    }
}

function updateMeasurementForPolygon(poly) {
    // 使用從 fix-fetch-and-measurements.js 補丁版本
    if (typeof window.updateMeasurementForPolygon_safe === 'function') {
        return window.updateMeasurementForPolygon_safe(poly);
    }
    
    if (!poly || !poly.centroid) return;
    
    poly.measurementData = {
        centroid: poly.centroid,
        distances: {}
    };
    
    // 計算到各個 checkbox 的距離
    const checkboxes = document.querySelectorAll('#panel input[type="checkbox"]');
    checkboxes.forEach(cb => {
        if (cb.checked) {
            const value = cb.value;
            const result = results.find(r => r.type === value);
            if (result && typeof haversine === 'function') {
                const dist = haversine(poly.centroid[0], poly.centroid[1], result.lat, result.lon);
                poly.measurementData.distances[value] = dist;
            }
        }
    });
    
    displayMeasurement(poly);
}

function displayMeasurement(poly) {
    const panel = document.getElementById('measurementPanel');
    const content = document.getElementById('measurementContent');
    
    if (!poly || !poly.measurementData) {
        panel.style.display = 'none';
        return;
    }
    
    let html = `<div style="font-weight:bold; margin-bottom:8px;">多邊形 #${poly.id} 形心量測</div>`;
    html += `<div style="margin-bottom:6px;"><b>形心座標:</b><br>
            Lat: ${poly.centroid[0].toFixed(6)}<br>
            Lon: ${poly.centroid[1].toFixed(6)}</div>`;
    
    html += `<div><b>距離量測 (m):</b><br>`;
    
    if (Object.keys(poly.measurementData.distances).length > 0) {
        Object.entries(poly.measurementData.distances).forEach(([key, dist]) => {
            const label = document.querySelector(`input[value="${key}"]`).parentElement.textContent.trim();
            html += `<div style="margin:4px 0; padding:4px; background:#f0f0f0; border-radius:4px;">
                    <strong>${label.substring(key.length)}</strong><br>
                    ${(dist/1000).toFixed(2)} km (${Math.round(dist)} m)
                    </div>`;
        });
    } else {
        html += '<div style="color:#999;">暫無量測數據</div>';
    }
    html += '</div>';
    
    content.innerHTML = html;
    panel.style.display = 'block';
}

function toggleMeasurementPanel() {
    const panel = document.getElementById('measurementPanel');
    panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
}

// ============================================================
// Haversine 距離計算
// ============================================================
function haversine(lat1, lon1, lat2, lon2) {
    const R = 6371000; // 地球半徑(m)
    const rad = Math.PI / 180;
    const dLat = (lat2 - lat1) * rad;
    const dLon = (lon2 - lon1) * rad;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
              Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * 
              Math.sin(dLon/2) * Math.sin(dLon/2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
    return R * c;
}

// ============================================================
// 查詢功能 - 使用 Overpass API (支援 CORS)
// ============================================================
async function runQuery() {
    const lat = parseFloat(latInput.value);
    const lon = parseFloat(lonInput.value);
    const radius = parseFloat(radiusInput.value);
    
    if (isNaN(lat) || isNaN(lon) || isNaN(radius)) {
        alert('請輸入有效的緯度、經度和半徑');
        return;
    }
    
    origin = [lat, lon];
    
    // 更新標記位置
    if (userMarker) map.removeLayer(userMarker);
    userMarker = L.marker(origin, {
        icon: L.divIcon({
            className: '',
            html: `<div style="width:18px;height:18px;background:#ef4444;border:3px solid #ffffff;border-radius:50%;box-shadow:0 0 0 2px rgba(239,68,68,0.35);"></div>`,
            iconSize: [18, 18],
            iconAnchor: [9, 9]
        }),
        interactive: false
    }).addTo(map);
    
    map.setView(origin, 15);
    
    // 禁用查詢按鈕
    queryBtn.disabled = true;
    queryBtn.innerText = '查詢中...';
    
    results = [];
    tbody.innerHTML = '';
    
    // 獲取選中的類型
    const checkboxes = document.querySelectorAll('#panel input[type="checkbox"]:checked');
    const types = Array.from(checkboxes).map(cb => cb.value);
    
    if (types.length === 0) {
        alert('請至少選擇一個類型');
        queryBtn.disabled = false;
        queryBtn.innerText = '查詢';
        return;
    }
    
    // 並行查詢各類型
    try {
        const promises = types.map(type => queryOverpassAPI(lat, lon, radius, type));
        const responses = await Promise.all(promises);
        
        responses.forEach((data, idx) => {
            if (data && data.elements && data.elements.length > 0) {
                data.elements.forEach(elem => {
                    if (elem.lat && elem.lon) {
                        const dist = haversine(lat, lon, elem.lat, elem.lon);
                        results.push({
                            name: elem.tags.name || '未命名',
                            type: types[idx],
                            lat: elem.lat,
                            lon: elem.lon,
                            distance: dist,
                            route: null
                        });
                    }
                });
            }
        });
        
        // 顯示結果
        displayResults();
        
        // 在地圖上標記結果
        drawResultMarkers();
        
    } catch (error) {
        console.error('[runQuery] 錯誤:', error);
        alert('查詢失敗: ' + error.message);
    } finally {
        queryBtn.disabled = false;
        queryBtn.innerText = '查詢';
    }
}

// ============================================================
// Overpass API 查詢
// ============================================================
async function queryOverpassAPI(lat, lon, radius, type) {
    const overpassUrl = 'https://overpass-api.de/api/interpreter';
    
    const tagMap = {
        'station': 'railway=station',
        'bus_stop': 'highway=bus_stop',
        'motorway_junction': 'motorway_junction=yes',
        'school': 'amenity=school',
        'marketplace': 'shop=supermarket|amenity=marketplace',
        'park': 'leisure=park',
        'tourism': 'tourism=attraction',
        'parking': 'amenity=parking',
        'bank': 'amenity=bank',
        'gas': 'man_made=gas_well',
        'cemetery': 'amenity=grave_yard',
        'waste_disposal': 'amenity=waste_disposal',
        'pollution': 'man_made=wastewater_plant'
    };
    
    const query = `
        [bbox:${lat-radius/111000},${lon-radius/111000/Math.cos(lat*Math.PI/180)},${lat+radius/111000},${lon+radius/111000/Math.cos(lat*Math.PI/180)}];
        (
            node[${tagMap[type]}];
            way[${tagMap[type]}];
            relation[${tagMap[type]}];
        );
        out center;
    `;
    
    try {
        const response = await safeFetch(overpassUrl, {
            method: 'POST',
            body: query
        });
        
        return typeof response === 'string' ? JSON.parse(response) : response;
    } catch (error) {
        console.error(`[queryOverpassAPI] 查詢 ${type} 失敗:`, error);
        return { elements: [] };
    }
}

function displayResults() {
    if (results.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#999;">未找到結果</td></tr>';
        return;
    }
    
    tbody.innerHTML = '';
    results.forEach((item, idx) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.name}</td>
            <td>${item.type}</td>
            <td>${(item.distance/1000).toFixed(2)} km</td>
            <td>${item.route ? (item.route/1000).toFixed(2) + ' km' : '-'}</td>
        `;
        tbody.appendChild(row);
    });
}

function drawResultMarkers() {
    markerLayer.clearLayers();
    results.forEach(item => {
        L.circleMarker([item.lat, item.lon], {
            radius: 5,
            fillColor: '#2563eb',
            color: '#1d4ed8',
            weight: 2,
            opacity: 0.8,
            fillOpacity: 0.6
        }).addTo(markerLayer).bindPopup(`
            <div style="font-size:12px;">
                <b>${item.name}</b><br>
                類型: ${item.type}<br>
                距離: ${(item.distance/1000).toFixed(2)} km
            </div>
        `);
    });
}

function sortResults(field) {
    sortState[field] = !sortState[field];
    results.sort((a, b) => {
        const aVal = a[field];
        const bVal = b[field];
        return sortState[field] ? aVal - bVal : bVal - aVal;
    });
    displayResults();
}

function exportExcel() {
    if (results.length === 0) {
        alert('沒有結果可匯出');
        return;
    }
    
    const data = results.map(r => ({
        '名稱': r.name,
        '類型': r.type,
        '直線距離(m)': Math.round(r.distance),
        '路線距離(m)': r.route ? Math.round(r.route) : ''
    }));
    
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '查詢結果');
    XLSX.writeFile(wb, '地價因素查詢結果.xlsx');
}

function clearRoute() {
    results.forEach(r => r.route = null);
    if (routeLayer) map.removeLayer(routeLayer);
    routeInfo.innerHTML = '路線距離：尚未計算';
    displayResults();
}

function getCurrentLocation() {
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(
            pos => {
                const lat = pos.coords.latitude;
                const lon = pos.coords.longitude;
                latInput.value = lat.toFixed(6);
                lonInput.value = lon.toFixed(6);
                alert('已取得位置: ' + lat.toFixed(6) + ', ' + lon.toFixed(6));
            },
            err => alert('無法取得位置: ' + err.message)
        );
    } else {
        alert('瀏覽器不支援地理定位');
    }
}

function analyzeSlope() {
    alert('坡度分析功能開發中...');
}

function exportAllPolygonsSHP() {
    if (polygons.length === 0) {
        alert("目前沒有多邊形可供匯出！");
        return;
    }
    
    try {
        let features = polygons.map(poly => poly.toGeoJSON());
        let geojsonObj = {
            type: "FeatureCollection",
            features: features
        };
        
        let dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(geojsonObj));
        let link = document.createElement('a');
        link.href = dataStr;
        link.download = "polygons.geojson";
        link.click();
        
        alert('已匯出 GeoJSON 檔案');
    } catch (error) {
        console.error('[exportAllPolygonsSHP] 錯誤:', error);
        alert('匯出失敗: ' + error.message);
    }
}

function importPolygonSHP(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    if (file.name.endsWith('.zip')) {
        // SHP ZIP 處理
        JSZip.loadAsync(file).then(zip => {
            // 基礎 SHP 解析 (可擴展)
            alert('SHP 匯入功能開發中');
        });
    } else if (file.name.endsWith('.geojson') || file.name.endsWith('.json')) {
        // GeoJSON 處理
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const geojson = JSON.parse(e.target.result);
                if (geojson.type === 'FeatureCollection') {
                    let newId = polygons.length > 0 ? Math.max(...polygons.map(p => p.id)) + 1 : 1;
                    geojson.features.forEach(feature => {
                        if (feature.geometry.type === 'Polygon') {
                            let poly = Polygon.fromGeoJSON(feature, newId++);
                            polygons.push(poly);
                            poly.render();
                        }
                    });
                    updatePolygonList();
                    alert('已匯入 ' + geojson.features.length + ' 個多邊形');
                }
            } catch (error) {
                alert('GeoJSON 匯入失敗: ' + error.message);
            }
        };
        reader.readAsText(file);
    }
}

// 地圖點擊事件 - 繪製多邊形
map.on('click', function(e) {
    if (isDrawingPolygon) {
        currentDrawPoints.push([e.latlng.lat, e.latlng.lng]);
        
        // 顯示標記
        L.marker([e.latlng.lat, e.latlng.lng], {
            icon: L.divIcon({
                className: '',
                html: `<div style="width:10px;height:10px;background:#ef4444;border-radius:50%;"></div>`,
                iconSize: [10, 10],
                iconAnchor: [5, 5]
            })
        }).addTo(drawMarkers);
        
        // 更新預覽線
        if (previewPolyline) map.removeLayer(previewPolyline);
        if (currentDrawPoints.length > 1) {
            previewPolyline = L.polyline(currentDrawPoints, {
                color: '#2563eb',
                weight: 2,
                dashArray: '5,5'
            }).addTo(map);
        }
    }
});

console.log('頁面初始化完成');
</script>

</body>
</html>
