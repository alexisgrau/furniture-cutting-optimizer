// index.js - Point d'entrée du programme
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

// ============= CONFIGURATION =============
const CONFIG = {
    inputFile: 'input.xlsx',
    outputDir: 'output',
    boards: [
        {
            width: 2000,
            height: 500,
            thickness: 16,
            price: 10.9
        }
    ],
    kerf: 3,              // Perte par découpe (mm)
    margin: 4            // Marge de sécurité (mm)
};

// ============= INTERFACE UTILISATEUR =============
const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

function askQuestion(query) {
    return new Promise(resolve => rl.question(query, resolve));
}

// ============= LECTURE EXCEL =============
function readExcelFile(filename) {
    const workbook = XLSX.readFile(filename);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const data = [];
    const range = XLSX.utils.decode_range(sheet['!ref']);

    // Commencer à la ligne 2 (après l'en-tête)
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
        const nameCell = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
        const dim1Cell = sheet[XLSX.utils.encode_cell({ r: row, c: 1 })];
        const dim2Cell = sheet[XLSX.utils.encode_cell({ r: row, c: 3 })];
        const thicknessCell = sheet[XLSX.utils.encode_cell({ r: row, c: 4 })];
        const quantityCell = sheet[XLSX.utils.encode_cell({ r: row, c: 5 })];

        if (!nameCell || !dim1Cell || !dim2Cell){ continue; }

        const name = nameCell.v;
        const width = parseFloat(String(dim1Cell.v).replace(/[^\d.]/g, ''));
        const height = parseFloat(String(dim2Cell.v).replace(/[^\d.]/g, ''));
        const thickness = thicknessCell ? parseFloat(String(thicknessCell.v).replace(/[^\d.]/g, '')) : 16;
        const quantity = quantityCell ? parseInt(quantityCell.v) : 1;

        if (name && !isNaN(width) && !isNaN(height)) {
            for (let i = 0; i < quantity; i++) {
                data.push({
                    name: quantity > 1 ? `${name} (${i + 1}/${quantity})` : name,
                    width,
                    height,
                    thickness,
                    area: width * height,
                    originalQuantity: quantity
                });
            }
        }
    }

    return data;
}

// ============= AFFICHAGE DES DONNÉES =============
function displayDataPreview(pieces) {
    console.log('\n' + '='.repeat(80));
    console.log('APERÇU DES DONNÉES CHARGÉES (max 10)');
    console.log('='.repeat(80));

    console.log('-'.repeat(80));
    console.log(
        'N°'.padEnd(4) +
        'Nom'.padEnd(30) +
        'Largeur'.padEnd(12) +
        'Hauteur'.padEnd(12) +
        'Épaisseur'.padEnd(12) +
        'Qté'
    );
    console.log('-'.repeat(80));

    const displayedNames = new Set();
    let count = 0;

    for (const piece of pieces) {
        if (count >= 10) break;

        const baseName = piece.name.replace(/ \(\d+\/\d+\)$/, '');
        if (!displayedNames.has(baseName)) {
            displayedNames.add(baseName);
            console.log(
                `${(count + 1).toString().padEnd(4)}` +
                `${baseName.substring(0, 28).padEnd(30)}` +
                `${piece.width.toString().padEnd(12)}` +
                `${piece.height.toString().padEnd(12)}` +
                `${piece.thickness.toString().padEnd(12)}` +
                `${piece.originalQuantity || 1}`
            );
            count++;
        }
    }

    if (pieces.length > 10) {
        console.log(`... et ${pieces.length - count} autres éléments`);
    }

    console.log('-'.repeat(80));
    console.log(`\nTOTAL: ${pieces.length} pièces à découper\n`);
}



// ============= FILTRAGE PAR ÉPAISSEUR =============
function filterPiecesByAvailableBoards(pieces, boards) {
    const availableThicknesses = boards.map(b => b.thickness);
    const filtered = pieces.filter(p => availableThicknesses.includes(p.thickness));
    const excluded = pieces.filter(p => !availableThicknesses.includes(p.thickness));

    if (excluded.length > 0) {
        console.log('\nATTENTION: Pièces non prises en compte (épaisseur non configurée):');
        const byThickness = {};
        excluded.forEach(p => {
            if (!byThickness[p.thickness]) byThickness[p.thickness] = 0;
            byThickness[p.thickness]++;
        });
        Object.keys(byThickness).forEach(t => {
            console.log(`   - ${byThickness[t]} pièces de ${t}mm d'épaisseur`);
        });
        console.log('\nAjoutez ces épaisseurs dans CONFIG.boards si nécessaire\n');
    }

    return filtered;
}

// ============= ALGORITHME D'OPTIMISATION =============
class CuttingOptimizer {
    constructor(config) {
        this.config = config;
        this.boards = [];
    }

    optimize(pieces) {
        // Grouper par épaisseur
        const byThickness = {};
        pieces.forEach(piece => {
            if (!byThickness[piece.thickness]) {
                byThickness[piece.thickness] = [];
            }
            byThickness[piece.thickness].push(piece);
        });

        // Optimiser chaque groupe avec la config de planche appropriée
        Object.keys(byThickness).forEach(thickness => {
            const piecesForThickness = byThickness[thickness];
            const boardConfig = this.config.boards.find(b => b.thickness == thickness);

            if (!boardConfig) {
                console.warn(`   Pas de configuration de planche pour épaisseur ${thickness}mm`);
                return;
            }

            // Trier par aire décroissante
            piecesForThickness.sort((a, b) => b.area - a.area);

            this.packPieces(piecesForThickness, parseFloat(thickness), boardConfig);
        });

        return this.boards;
    }

    packPieces(pieces, thickness, boardConfig) {
        const kerf = this.config.kerf;

        for (const piece of pieces) {
            let placed = false;

            // Essayer de placer sur une planche existante
            for (const board of this.boards) {
                if (board.thickness !== thickness) continue;

                if (this.tryPlacePiece(board, piece, kerf)) {
                    placed = true;
                    break;
                }
            }

            // Créer une nouvelle planche si nécessaire
            if (!placed) {
                const newBoard = {
                    id: this.boards.length + 1,
                    thickness,
                    width: boardConfig.width,
                    height: boardConfig.height,
                    price: boardConfig.price,
                    pieces: [],
                    usedSpace: []
                };

                if (this.tryPlacePiece(newBoard, piece, kerf)) {
                    this.boards.push(newBoard);
                } else {
                    console.warn(`   Pièce trop grande: ${piece.name} (${piece.width}x${piece.height}mm) pour planche ${boardConfig.width}x${boardConfig.height}mm`);
                }
            }
        }
    }

    tryPlacePiece(board, piece, kerf) {
        const margin = this.config.margin;

        // Essayer les deux orientations
        const orientations = [
            { w: piece.width, h: piece.height, rotated: false },
            { w: piece.height, h: piece.width, rotated: true }
        ];

        for (const orient of orientations) {
            const w = orient.w + kerf;
            const h = orient.h + kerf;

            if (w > board.width || h > board.height) continue;

            // Algorithme de placement: essayer chaque position
            for (let y = margin; y <= board.height - h - margin; y += 10) {
                for (let x = margin; x <= board.width - w - margin; x += 10) {
                    if (this.canPlaceAt(board, x, y, w, h)) {
                        board.pieces.push({
                            ...piece,
                            x,
                            y,
                            placedWidth: orient.w,
                            placedHeight: orient.h,
                            rotated: orient.rotated
                        });
                        board.usedSpace.push({ x, y, w, h });
                        return true;
                    }
                }
            }
        }

        return false;
    }

    canPlaceAt(board, x, y, w, h) {
        for (const used of board.usedSpace) {
            if (!(x + w <= used.x || x >= used.x + used.w ||
                y + h <= used.y || y >= used.y + used.h)) {
                return false;
            }
        }
        return true;
    }
}

// ============= GÉNÉRATION SVG =============
function generateSVG(board, config) {
    const scale = 0.3; // Échelle pour l'affichage
    const svgWidth = board.width * scale;
    const svgHeight = board.height * scale;

    let svg = `<svg width="${svgWidth}" height="${svgHeight}" xmlns="http://www.w3.org/2000/svg">`;

    // Fond de la planche
    svg += `<rect width="${svgWidth}" height="${svgHeight}" fill="#f5f5f5" stroke="#333" stroke-width="2"/>`;

    // Dessiner chaque pièce
    const colors = ['#FFB6C1', '#87CEEB', '#98FB98', '#FFD700', '#DDA0DD', '#F0E68C', '#FFB347', '#B0E0E6'];

    board.pieces.forEach((piece, idx) => {
        const x = piece.x * scale;
        const y = piece.y * scale;
        const w = piece.placedWidth * scale;
        const h = piece.placedHeight * scale;
        const color = colors[idx % colors.length];

        // Rectangle de la pièce
        svg += `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="${color}" stroke="#333" stroke-width="1" opacity="0.8"/>`;

        // Texte avec le nom
        const fontSize = Math.min(12, w / 10, h / 5);
        const textX = x + w / 2;
        const textY = y + h / 2;

        // Nom de la pièce
        svg += `<text x="${textX}" y="${textY - 5}" text-anchor="middle" font-size="${fontSize}" font-weight="bold" fill="#000">${piece.name}</text>`;

        // Dimensions
        svg += `<text x="${textX}" y="${textY + fontSize}" text-anchor="middle" font-size="${fontSize * 0.8}" fill="#333">${piece.placedWidth}×${piece.placedHeight}mm</text>`;

        // Indicateur de rotation
        if (piece.rotated) {
            svg += `<text x="${textX}" y="${textY + fontSize * 2}" text-anchor="middle" font-size="${fontSize * 0.7}" fill="#d00">↻ Pivoté</text>`;
        }
    });

    svg += '</svg>';
    return svg;
}

// ============= GÉNÉRATION HTML =============
function generateHTML(boards, config, stats) {
    let html = `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Plan de Découpe Optimisé</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #fff; }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        h2 { color: #34495e; margin-top: 0; }
        .summary { background: #ecf0f1; padding: 20px; border-radius: 8px; margin: 20px 0; }
        .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; }
        .stat { background: white; padding: 15px; border-radius: 5px; border-left: 4px solid #3498db; }
        .stat-label { font-size: 0.9em; color: #7f8c8d; }
        .stat-value { font-size: 1.8em; font-weight: bold; color: #2c3e50; }
        .board { margin: 30px 0; padding: 20px; border: 2px solid #bdc3c7; border-radius: 8px; background: white; }
        .board-header { background: #3498db; color: white; padding: 10px; border-radius: 5px; margin-bottom: 15px; }
        .piece-list { margin-top: 15px; }
        .piece-item { padding: 8px; margin: 5px 0; background: #f8f9fa; }
        svg { border: 1px solid #ddd; border-radius: 5px; margin: 10px 0; background: white; }
        .print-button { background: #3498db; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 16px; margin: 10px 0; }
        .print-button:hover { background: #2980b9; }
        .shopping-list { background: #fff3cd; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #ffc107; }
        @media print {
            .print-button { display: none; }
            .board { page-break-after: always; }
        }
    </style>
</head>
<body>
    <h1>Plan de Découpe Optimisé</h1>
    <button class="print-button" onclick="window.print()">🖨️ Imprimer</button>
    
    <div class="shopping-list">
        <h2>Liste de courses</h2>
        ${stats.boardsByType.map(bt => `
            <div style="margin: 10px 0;">
                <strong>${bt.count}x</strong> Planches ${bt.width}×${bt.height}mm - Épaisseur ${bt.thickness}mm 
                @ ${bt.price}€/unité = <strong>${(bt.count * bt.price).toFixed(2)}€</strong>
            </div>
        `).join('')}
        <div style="margin-top: 15px; font-size: 1.2em; font-weight: bold; color: #2c3e50;">
            Total: ${stats.totalCost.toFixed(2)} €
        </div>
    </div>
    
    <div class="summary">
        <h2>Résumé</h2>
        <div class="summary-grid">
            <div class="stat">
                <div class="stat-label">Nombre de planches</div>
                <div class="stat-value">${stats.totalBoards}</div>
            </div>
            <div class="stat">
                <div class="stat-label">Coût total</div>
                <div class="stat-value">${stats.totalCost.toFixed(2)} €</div>
            </div>
            <div class="stat">
                <div class="stat-label">Pièces à découper</div>
                <div class="stat-value">${stats.totalPieces}</div>
            </div>
            <div class="stat">
                <div class="stat-label">Taux d'utilisation</div>
                <div class="stat-value">${stats.efficiency.toFixed(1)}%</div>
            </div>
        </div>
    </div>`;

    boards.forEach((board, idx) => {
        const usedArea = board.pieces.reduce((sum, p) => sum + (p.placedWidth * p.placedHeight), 0);
        const boardArea = board.width * board.height;
        const efficiency = (usedArea / boardArea * 100).toFixed(1);

        html += `
    <div class="board">
        <div class="board-header">
            <h3>Planche #${board.id} - ${board.width}×${board.height}mm - Épaisseur: ${board.thickness}mm (Utilisation: ${efficiency}%)</h3>
        </div>
        ${generateSVG(board, config)}
        
        <div class="piece-list">
            <h4>Liste des découpes (${board.pieces.length} pièces):</h4>
            ${board.pieces.map((p, i) => `
                <div class="piece-item">
                    <strong>${i + 1}. ${p.name}</strong> - ${p.placedWidth} × ${p.placedHeight} mm
                    ${p.rotated ? '↻ <em>(Pivoté à 90°)</em>' : ''}
                    - Position: X=${p.x}mm, Y=${p.y}mm
                </div>
            `).join('')}
        </div>
    </div>`;
    });

    html += `
</body>
</html>`;

    return html;
}

// ============= PROGRAMME PRINCIPAL =============
async function main() {
    console.log('Démarrage de l\'optimiseur de découpe...\n');

    // Vérifier le fichier d'entrée
    if (!fs.existsSync(CONFIG.inputFile)) {
        console.error(`Fichier non trouvé: ${CONFIG.inputFile}`);
        console.log('Placez votre fichier Excel dans le même dossier et nommez-le "input.xlsx"');
        rl.close();
        return;
    }

    // Afficher la configuration
    console.log('Configuration des planches disponibles:');
    CONFIG.boards.forEach((board, idx) => {
        console.log(`   ${idx + 1}. ${board.width}×${board.height}mm - Épaisseur ${board.thickness}mm - ${board.price}€`);
    });

    // Lire les données
    console.log('\nLecture du fichier Excel...');
    const allPieces = readExcelFile(CONFIG.inputFile);

    // Afficher l'aperçu
    displayDataPreview(allPieces);

    // Demander confirmation
    const answer = await askQuestion('Voulez-vous continuer avec ces données ? (Y/N): ');

    if (answer.trim().toLowerCase() !== 'y') {
        console.log('❌ Opération annulée.');
        rl.close();
        return;
    }

    // Filtrer les pièces selon les épaisseurs disponibles
    const pieces = filterPiecesByAvailableBoards(allPieces, CONFIG.boards);

    if (pieces.length === 0) {
        console.error('❌ Aucune pièce ne correspond aux épaisseurs configurées!');
        rl.close();
        return;
    }

    console.log(`\n${pieces.length} pièces retenues pour l'optimisation\n`);

    // Optimiser
    console.log('Optimisation du placement...');
    const optimizer = new CuttingOptimizer(CONFIG);
    const boards = optimizer.optimize(pieces);
    console.log(`   - ${boards.length} planches nécessaires\n`);

    // Calculer les statistiques
    const boardsByType = {};
    boards.forEach(b => {
        const key = `${b.width}x${b.height}x${b.thickness}`;
        if (!boardsByType[key]) {
            boardsByType[key] = {
                width: b.width,
                height: b.height,
                thickness: b.thickness,
                price: b.price,
                count: 0
            };
        }
        boardsByType[key].count++;
    });

    const stats = {
        totalBoards: boards.length,
        totalPieces: pieces.length,
        totalCost: boards.reduce((sum, b) => sum + b.price, 0),
        boardsByType: Object.values(boardsByType),
        efficiency: 0
    };

    const totalUsedArea = boards.reduce((sum, b) => {
        return sum + b.pieces.reduce((s, p) => s + (p.placedWidth * p.placedHeight), 0);
    }, 0);
    const totalBoardArea = boards.reduce((sum, b) => sum + (b.width * b.height), 0);
    stats.efficiency = totalBoardArea > 0 ? (totalUsedArea / totalBoardArea * 100) : 0;

    // Générer le HTML
    console.log('Génération du rapport HTML...');
    const html = generateHTML(boards, CONFIG, stats);

    // Créer le dossier de sortie
    if (!fs.existsSync(CONFIG.outputDir)) {
        fs.mkdirSync(CONFIG.outputDir);
    }

    // Écrire le fichier
    const outputPath = path.join(CONFIG.outputDir, 'results.html');
    fs.writeFileSync(outputPath, html);

    console.log('\n' + '='.repeat(80));
    console.log('TERMINÉ!');
    console.log('='.repeat(80));
    console.log('\nRésumé:');
    console.log(`   - Planches nécessaires: ${stats.totalBoards}`);
    stats.boardsByType.forEach(bt => {
        console.log(`     • ${bt.count}x ${bt.width}×${bt.height}mm (${bt.thickness}mm) @ ${bt.price}€ = ${(bt.count * bt.price).toFixed(2)}€`);
    });
    console.log(`   - Coût total: ${stats.totalCost.toFixed(2)} €`);
    console.log(`   - Taux d'utilisation: ${stats.efficiency.toFixed(1)}%`);
    console.log(`\n Rapport généré: ${outputPath}`);
    console.log(`\n Ouvrez le fichier avec votre navigateur pour voir les plans de découpe!`);
    console.log('\n' + '='.repeat(80) + '\n');

    rl.close();
}

main().catch(error => {
    console.error('❌ Erreur:', error.message);
    console.error(error.stack);
    rl.close();
});