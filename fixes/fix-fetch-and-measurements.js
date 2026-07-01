// fixes/fix-fetch-and-measurements.js
// This file provides safeFetch and tolerant measurement helpers.
// To apply immediately, add a script tag in index.html pointing to this file:
// <script src="fixes/fix-fetch-and-measurements.js"></script>

(function(){
  console.log('[patch] fix-fetch-and-measurements loaded');

  async function safeFetch(url, options={}){
    try{
      const resp = await fetch(url, Object.assign({mode:'cors'}, options));
      if (!resp.ok) {
        let text = '';
        try{ text = await resp.text(); } catch(e){ text = '<no body>'; }
        throw new Error(`HTTP ${resp.status} ${resp.statusText} - ${text}`);
      }
      const ct = resp.headers.get('content-type') || '';
      if (ct.includes('application/json')) return await resp.json();
      return await resp.text();
    } catch(err){
      console.error('[safeFetch] failed', url, err);
      throw err;
    }
  }

  // Expose globally for other scripts
  window.safeFetch = safeFetch;

  // If runQuery exists, wrap it to show better errors
  if (typeof window.runQuery === 'function'){
    const originalRunQuery = window.runQuery;
    window.runQuery = async function(...args){
      try{
        const btn = document.getElementById('query-btn');
        if (btn) btn.disabled = true;
        return await originalRunQuery.apply(this, args);
      } catch(err){
        console.error('[patched runQuery] error', err);
        alert('查詢失敗: ' + (err && err.message ? err.message : err));
      } finally{
        const btn = document.getElementById('query-btn');
        if (btn) btn.disabled = false;
      }
    };
  }

  // Provide tolerant updateMeasurementForPolygon when results[] may be empty
  window.updateMeasurementForPolygon_safe = function(poly){
    if (!poly || !poly.centroid) return;
    poly.measurementData = { centroid: poly.centroid, distances: {} };
    try{
      const checkboxes = document.querySelectorAll('#panel input[type="checkbox"]');
      checkboxes.forEach(cb => {
        if (cb.checked){
          const value = cb.value;
          const result = (window.results || []).find(r => r.type === value);
          if (result && typeof result.lat === 'number' && typeof result.lon === 'number'){
            const dist = (typeof window.haversine === 'function') ? window.haversine(poly.centroid[0], poly.centroid[1], result.lat, result.lon) : null;
            poly.measurementData.distances[value] = dist === null ? null : dist;
          } else {
            // mark as null to indicate no data
            poly.measurementData.distances[value] = null;
          }
        }
      });
    } catch(e){
      console.error('[updateMeasurementForPolygon_safe] error', e);
    }
    // call original displayMeasurement if exists, otherwise use built-in safe display
    if (typeof window.displayMeasurement === 'function'){
      try{ window.displayMeasurement(poly); } catch(e){ console.error('[patched displayMeasurement] failed', e); }
    } else {
      // very small fallback
      const panel = document.getElementById('measurementPanel');
      const content = document.getElementById('measurementContent');
      let html = `<div style="font-weight:bold; margin-bottom:8px;">多邊形 #${poly.id} 形心量測（備援顯示）</div>`;
      html += `<div><b>形心:</b>${poly.centroid[0].toFixed(6)}, ${poly.centroid[1].toFixed(6)}</div>`;
      html += '<div><b>距離:</b><br/>';
      Object.entries(poly.measurementData.distances).forEach(([k,v]) => {
        html += `<div style="padding:4px;background:#f0f0f0;margin:4px 0;border-radius:4px;">${k}: ${v===null? '無資料' : Math.round(v) + ' m'}</div>`;
      });
      html += '</div>';
      content.innerHTML = html;
      panel.style.display = 'block';
    }
  };

  // If original updateMeasurementForPolygon exists, keep copy and override to safer version
  if (typeof window.updateMeasurementForPolygon === 'function'){
    window._orig_updateMeasurementForPolygon = window.updateMeasurementForPolygon;
    window.updateMeasurementForPolygon = window.updateMeasurementForPolygon_safe;
    console.log('[patch] updateMeasurementForPolygon overridden');
  } else {
    window.updateMeasurementForPolygon = window.updateMeasurementForPolygon_safe;
    console.log('[patch] updateMeasurementForPolygon added');
  }

  // Provide a small helper to patch displayMeasurement to handle null distances
  if (typeof window.displayMeasurement === 'function'){
    const origDisplay = window.displayMeasurement;
    window.displayMeasurement = function(poly){
      try{
        // call original, but guard accesses
        return origDisplay.call(this, poly);
      } catch(e){
        console.error('[patched displayMeasurement] original failed, using fallback', e);
        return window.updateMeasurementForPolygon_safe(poly);
      }
    };
    console.log('[patch] displayMeasurement wrapped');
  }

})();
