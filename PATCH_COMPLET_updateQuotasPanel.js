// =====================================================================
// FONCTION COMPLÈTE updateQuotasPanel - À COPIER DANS ConsolePilotageV3.html
// Remplace la fonction existante (lignes 1702-1862 environ)
// =====================================================================

function updateQuotasPanel(stats, structData) {
    // Calculer les quotas totaux depuis la structure
    const quotas = { lv2: {}, options: {} };
    let totalPlaces = 0;
    
    if (structData && structData.classes) {
        structData.classes.forEach(cls => {
            // Additionner les capacités totales
            totalPlaces += parseInt(cls.capacity) || 0;
            
            // Additionner les quotas de chaque classe
            Object.entries(cls.quotas || {}).forEach(([key, val]) => {
                const value = parseInt(val) || 0;
                if (structData.lv2 && structData.lv2.includes(key)) {
                    quotas.lv2[key] = (quotas.lv2[key] || 0) + value;
                } else if (structData.options && structData.options.includes(key)) {
                    quotas.options[key] = (quotas.options[key] || 0) + value;
                }
            });
        });
    }
    
    // === ALERTE GLOBALE PLACES vs ÉLÈVES (Glance Value) ===
    const totalEleves = stats.effectifs.total;
    const placesEl = document.getElementById('metric-places');
    if (placesEl) {
        const diff = totalPlaces - totalEleves;
        let color, bgColor, icon;

        if (diff < 0) {
            color = '#ef4444';
            bgColor = 'rgba(239, 68, 68, 0.1)';
            icon = '⚠️ ';
        } else if (diff < 5) {
            color = '#eab308';
            bgColor = 'rgba(234, 179, 8, 0.1)';
            icon = '⚠️ ';
        } else {
            color = '#22c55e';
            bgColor = 'rgba(34, 197, 94, 0.1)';
            icon = '✓ ';
        }

        placesEl.innerHTML = `
            <div style="background:${bgColor};padding:10px;border-radius:8px;border-left:3px solid ${color}">
                <div style="font-size:11px;color:#64748b;margin-bottom:4px;">🪑 PLACES vs ÉLÈVES</div>
                <div style="font-size:16px;font-weight:900;color:${color}">
                    ${totalEleves} élèves / ${totalPlaces} places ${icon}${diff}
                </div>
            </div>`;
    }

    // === LV2 - Format ULTRA-CLAIR X/Y - COMPTAGE GLOBAL ===
    let lv2HTML = '';
    Object.entries(quotas.lv2).forEach(([lv2, quota]) => {
        const count = stats.global[lv2] || 0;  // ✅ TOUS les élèves (seuls + avec option)
        const isComplete = count === quota;
        const isOver = count > quota;
        let barColor = isOver ? '#ef4444' : (isComplete ? '#22c55e' : '#eab308');
        let barWidth = Math.min(100, (count / Math.max(quota, 1)) * 100);
        const validIcon = isComplete ? '<span style="color:#22c55e;font-size:16px;margin-left:8px">✓</span>' : '';
        
        lv2HTML += `<tr>
            <td style="padding:8px;font-weight:700;font-size:13px;color:#e2e8f0">${lv2}</td>
            <td style="text-align:right;padding:8px">
                <div style="display:flex;align-items:center;justify-content:flex-end;gap:8px">
                    <div style="background:#1e293b;border-radius:8px;padding:4px 12px;min-width:60px;text-align:center">
                        <span style="color:#94a3b8;font-size:11px;font-weight:600">${quota}</span>
                    </div>
                    <div style="background:#ef4444;border-radius:8px;padding:4px 12px;min-width:60px;text-align:center">
                        <span style="color:#fff;font-size:13px;font-weight:900">${count}/${quota}</span>
                    </div>
                    ${validIcon}
                </div>
                <div style="background:#1e293b;height:4px;border-radius:2px;margin-top:6px;overflow:hidden">
                    <div style="background:${barColor};height:100%;width:${barWidth}%;transition:all 0.3s"></div>
                </div>
            </td>
        </tr>`;
    });
    document.getElementById('quota-lv2').innerHTML = lv2HTML || '<tr><td style="color:#64748b">Aucune</td></tr>';

    // === OPTIONS - Format ULTRA-CLAIR X/Y - COMPTAGE GLOBAL ===
    let optHTML = '';
    Object.entries(quotas.options).forEach(([opt, quota]) => {
        const count = stats.global[opt] || 0;  // ✅ TOUS les élèves (seuls + avec LV2)
        const isComplete = count === quota;
        const isOver = count > quota;
        let barColor = isOver ? '#ef4444' : (isComplete ? '#22c55e' : '#eab308');
        let barWidth = Math.min(100, (count / Math.max(quota, 1)) * 100);
        const validIcon = isComplete ? '<span style="color:#22c55e;font-size:16px;margin-left:8px">✓</span>' : '';
        
        optHTML += `<tr>
            <td style="padding:8px;font-weight:700;font-size:13px;color:#e2e8f0">${opt}</td>
            <td style="text-align:right;padding:8px">
                <div style="display:flex;align-items:center;justify-content:flex-end;gap:8px">
                    <div style="background:#1e293b;border-radius:8px;padding:4px 12px;min-width:60px;text-align:center">
                        <span style="color:#94a3b8;font-size:11px;font-weight:600">${quota}</span>
                    </div>
                    <div style="background:#ef4444;border-radius:8px;padding:4px 12px;min-width:60px;text-align:center">
                        <span style="color:#fff;font-size:13px;font-weight:900">${count}/${quota}</span>
                    </div>
                    ${validIcon}
                </div>
                <div style="background:#1e293b;height:4px;border-radius:2px;margin-top:6px;overflow:hidden">
                    <div style="background:${barColor};height:100%;width:${barWidth}%;transition:all 0.3s"></div>
                </div>
            </td>
        </tr>`;
    });
    document.getElementById('quota-options').innerHTML = optHTML || '<tr><td style="color:#64748b">Aucune</td></tr>';

    // === PROFILS DOUBLES - Toutes combinaisons LV2 × OPTIONS ===
    let comboHTML = '';
    const lv2Keys = Object.keys(quotas.lv2);
    const optKeys = Object.keys(quotas.options);
    
    const lv2KeysUpper = lv2Keys.map(k => k.toUpperCase());
    lv2Keys.forEach(lv2 => {
        optKeys.forEach(opt => {
            const comboKey = `${lv2} + ${opt}`;
            const count = stats.combos[comboKey] || 0;
            // CORRECTION : On additionne les places compatibles classe par classe
            let quotaMax = 0;
            (structData?.classes || []).forEach(cls => {
                const lv2Upper = lv2.toUpperCase();
                const optUpper = opt.toUpperCase();
                const quotasUpper = {};
                Object.keys(cls.quotas || {}).forEach(k => {
                    quotasUpper[k.toUpperCase()] = cls.quotas[k];
                });

                const qLv2 = parseInt(quotasUpper[lv2Upper] || 0);
                const qOpt = parseInt(quotasUpper[optUpper] || 0);
                if (qLv2 <= 0 || qOpt <= 0) return;

                // Classe incompatible si une autre LV2 couvre l'option presque totalement
                const hasConflictingLv2 = lv2KeysUpper.some(otherLv2 => {
                    if (otherLv2 === lv2Upper) return false;
                    const qOther = parseInt(quotasUpper[otherLv2] || 0);
                    if (qOther <= 0) return false;
                    const ratio = qOpt > 0 ? Math.min(qOther, qOpt) / qOpt : 0;
                    return ratio >= 0.9; // parfait ou quasi-parfait
                });
                if (hasConflictingLv2) return;

                quotaMax += Math.min(qLv2, qOpt);
            });
            const isComplete = count === quotaMax && quotaMax > 0;
            let barColor = count > quotaMax ? '#ef4444' : (isComplete ? '#22c55e' : '#eab308');
            let barWidth = quotaMax > 0 ? Math.min(100, (count / quotaMax) * 100) : 0;
            const validIcon = isComplete ? '<span style="color:#22c55e;font-size:16px;margin-left:8px">✓</span>' : '';
            
            comboHTML += `<tr>
                <td style="padding:8px;font-weight:700;font-size:12px;color:#eab308">${lv2} + ${opt}</td>
                <td style="text-align:right;padding:8px">
                    <div style="display:flex;align-items:center;justify-content:flex-end;gap:8px">
                        <div style="background:#ef4444;border-radius:8px;padding:4px 12px;min-width:50px;text-align:center">
                            <span style="color:#fff;font-size:12px;font-weight:900">${count}/${quotaMax}</span>
                        </div>
                        ${validIcon}
                    </div>
                    <div style="background:#1e293b;height:3px;border-radius:2px;margin-top:4px;overflow:hidden">
                        <div style="background:${barColor};height:100%;width:${barWidth}%;transition:all 0.3s"></div>
                    </div>
                </td>
            </tr>`;
        });
    });
    
    if (comboHTML === '') {
        comboHTML = '<tr><td colspan="2" style="color:#64748b;padding:8px">Aucune combinaison</td></tr>';
    }
    document.getElementById('quota-combos').innerHTML = comboHTML;

    // ASSO
    document.getElementById('quota-asso').innerHTML = 
        `<b>${stats.asso.codes}</b> codes<br><b>${stats.asso.eleves}</b> élèves`;

    // DISSO
    let dissoHTML = `<b>${stats.disso.codes}</b> codes<br><b>${stats.disso.eleves}</b> élèves`;
    if (stats.disso.conflicts && stats.disso.conflicts.length > 0) {
        dissoHTML += '<br><span style="color:var(--danger);font-size:11px">⚠️ ' + stats.disso.conflicts.length + ' conflits</span>';
    }
    document.getElementById('quota-disso').innerHTML = dissoHTML;
}
