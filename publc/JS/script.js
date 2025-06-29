/*
 * Copyright 2025 PROPUL. Todos los derechos reservados.
 * Script para la Calculadora de Materiales Tablayeso.
 * Maneja la lógica de agregar ítems, calcular materiales y generar reportes.
 * Implementa el criterio de cálculo v2.0, con nombres específicos para Durock Calibre 20 y lógica de tornillos de 1".
 * Implementa lógica de cálculo de paneles con acumuladores fraccionarios/redondeados para áreas pequeñas/grandes (según criterio de imagen).
 * Implementa selección de tipo de panel por cara de muro y para cielos.
 * Implementa Múltiples Entradas de Medida (Segmentos) para Muros Y Cielos.
 * Ajusta el orden de las entradas.
 * Implementa cálculo de Angular de Lámina para cielos basado en el perímetro completo de segmentos.
 * Agrega opción para descontar metraje de Angular en cielos.
 * Agrega resumen de opciones del ítem padre dentro de cada bloque de segmento en el área del encabezado.
 * Agrega campo para Área de Trabajo en el cálculo general y lo incluye en reportes.
 * Añade manejo básico de errores en el cálculo.
 * Implementa funcionalidad para importar segmentos desde archivo Excel.
 * INTEGRACIÓN: Añade lógica para cálculo de CENEFAS.
 * INTEGRACIÓN: Implementa REGLA GENERAL para dimensiones < 1m.
 * INTEGRACIÓN: Unifica cálculo final de Fijaciones.
 */

document.addEventListener('DOMContentLoaded', () => {
    // Get references to key DOM elements using IDs defined in index.html
    const itemsContainer = document.getElementById('items-container');
    const addItemBtn = document.getElementById('add-item-btn');
    const calculateBtn = document.getElementById('calculate-btn');
    const resultsContent = document.getElementById('results-content');
    const downloadOptionsDiv = document.querySelector('.download-options'); // Uses a class defined in style.css
    const generatePdfBtn = document.getElementById('generate-pdf-btn');
    const generateExcelBtn = document.getElementById('generate-excel-btn');
    const workAreaInput = document.getElementById('work-area'); // Added reference for the new input


    let itemCounter = 0; // To give unique IDs to item blocks

    // Variables to store the last calculated state (needed for PDF/Excel)
    let lastCalculatedTotalMaterials = {}; // Stores final rounded totals for all materials
    let lastCalculatedItemsSpecs = []; // Specs of items included in calculation
    let lastErrorMessages = []; // Store errors as an array of strings
    let lastCalculatedWorkArea = ''; // Variable para almacenar el Área de Trabajo calculada


    // --- Constants ---
    const PANEL_RENDIMIENTO_M2 = 2.98; // m2 por panel (rendimiento estándar de un panel de 1.22m x 2.44m)
    const SMALL_AREA_THRESHOLD_M2 = 1.5; // Umbral para considerar un área "pequeña" (en m2 por cara/área total del SEGMENTO)
    const POSTE_LARGO_ESTANDAR = 3.66; // Largo máximo estándar del poste para cálculo de empalmes
    const CANAL_LARGO_ESTANDAR = 3.05; // Largo estándar del canal para cálculo
    const ANGULAR_LARGO_ESTANDAR = 2.44; // Largo estándar del angular para cálculo
    const EMPALME_POSTE_LONGITUD = 0.30; // Longitud de empalme de poste
    const EMPALME_ANGULAR_LONGITUD = 0.15; // Requisito de empalme angular
    const LONGITUD_EXTRA_PATA = 0.10; // Longitud extra por soporte (pata)
    const ESPACIAMIENTO_CANAL_LISTON = 0.40; // Espaciamiento estándar para Canal Listón
    const ESPACIAMIENTO_CANAL_SOPORTE = 0.90; // Espaciamiento estándar para Canal Soporte


    // Definición de tipos de panel permitidos (deben coincidir con las opciones en el HTML)
    // Agregamos "Durock" según la matriz
    const PANEL_TYPES = [
        "Normal",
        "Resistente a la Humedad",
        "Resistente al Fuego",
        "Alta Resistencia",
        "Exterior", // Usado en la matriz original JS
        "Durock" // Nuevo tipo de panel explícito de la matriz unificada
    ];

     // --- Helper Function: REGLA GENERAL PARA CÁLCULO DE METRAJES ---
     // Si una dimensión es menor a 1, la aproxima a 1 para el cálculo
     const applyGeneralRule = (dimension) => {
         const num = parseFloat(dimension);
         if (isNaN(num) || num <= 0) return 0; // Treat invalid or zero input as zero
         return num < 1 ? 1.0 : num; // Apply the rule
     };


     // --- Helper Function for Rounding Up Final Units (Applies per item material quantity, EXCEPT panels in accumulators) ---
    const roundUpFinalUnit = (num) => Math.ceil(num);

    // --- Helper Function to get display name for item type ---
    const getItemTypeName = (typeValue) => {
        switch (typeValue) {
            case 'muro': return 'Muro';
            case 'cielo': return 'Cielo Falso';
            case 'cenefa': return 'Cenefa'; // New item type
            default: return 'Ítem Desconocido';
        }
    };

     // Helper to map item type internal value to a more descriptive name for inputs and summaries
     const getItemTypeDescription = (typeValue) => {
         switch (typeValue) {
             case 'muro': return 'Muro';
             case 'cielo': return 'Cielo Falso';
             case 'cenefa': return 'Cenefa'; // New item type
             default: return 'Ítem';
         }
     };

    // --- Helper Function to get the unit for a given material name ---
    const getMaterialUnit = (materialName) => {
         // Map specific names to units based on the new criterion
        // Material names can now include panel types, e.g., "Paneles de Normal"
        if (materialName.startsWith('Paneles de ')) return 'Und';

        switch (materialName) {
            case 'Postes': return 'Und';
            case 'Postes Calibre 20': return 'Und';
            case 'Canales': return 'Und';
            case 'Canales Calibre 20': return 'Und';
            case 'Pasta': return 'Caja';
            case 'Cinta de Papel': return 'm';
            case 'Lija Grano 120': return 'Pliego';
            case 'Clavos con Roldana': return 'Und';
            case 'Fulminantes': return 'Und';
            case 'Tornillos de 1" punta fina': return 'Und';
            case 'Tornillos de 1/2" punta fina': return 'Und';
            case 'Canal Listón': return 'Und';
            case 'Canal Soporte': return 'Und';
            case 'Angular de Lámina': return 'Und';
            case 'Tornillos de 1" punta broca': return 'Und';
            case 'Tornillos de 1/2" punta broca': return 'Und';
            case 'Patas': return 'Und';
            case 'Canal Listón (para cuelgue)': return 'Und';
            case 'Basecoat': return 'Saco';
            case 'Cinta malla': return 'm';
            // New materials for Cenefas structure (adjust units if necessary based on common packaging)
            case 'Canal Listón (Cenefa Horizontal)': return 'Und'; // Assuming pieces of 3.66m
            case 'Canal Listón (Cenefa Vertical)': return 'Und'; // Assuming pieces of 3.66m
            case 'Angular de Lámina (Cenefa)': return 'Und'; // Assuming pieces of 2.44m

            default: return 'Und'; // Default unit if not specified
        }
    };

    // Helper function to get the associated finishing materials based on panel type
    const getFinishingMaterials = (panelType) => {
         const finishing = {};
         // Associate finishing materials based on the panel type name or a category derived from it
         // Updated to include "Durock" as a separate type
         if (panelType === 'Normal' || panelType === 'Resistente a la Humedad' || panelType === 'Resistente al Fuego' || panelType === 'Alta Resistencia') {
             finishing['Pasta'] = 0;
             finishing['Cinta de Papel'] = 0;
             finishing['Lija Grano 120'] = 0;
             finishing['Tornillos de 1" punta fina'] = 0; // Yeso type screws for panel attachment
             finishing['Tornillos de 1/2" punta fina'] = 0; // Yeso type screws for structure
         } else if (panelType === 'Exterior' || panelType === 'Durock') { // Group Exterior and Durock for finishing
             finishing['Basecoat'] = 0;
             finishing['Cinta malla'] = 0;
             finishing['Tornillos de 1" punta broca'] = 0; // Durock type screws for panel attachment
             finishing['Tornillos de 1/2" punta broca'] = 0; // Durock type screws for structure
         }
         return finishing;
     };

    // --- Function to Populate Panel Type Selects ---
    const populatePanelTypes = (selectElement, selectedValue = 'Normal') => {
        selectElement.innerHTML = ''; // Clear existing options
        PANEL_TYPES.forEach(type => {
            const option = document.createElement('option');
            option.value = type;
            option.textContent = type;
            if (type === selectedValue) {
                option.selected = true;
            }
            selectElement.appendChild(option);
        });
    };

    // --- Function to update the summary details displayed within a segment block ---
    // This function reads the parent item's options and updates the summary div inside the segment.
    const updateSegmentItemSummary = (segmentBlock) => {
        // Find the parent item block from the segment block
        const itemBlock = segmentBlock.closest('.item-block');
        if (!itemBlock) {
            console.error("Could not find parent item block for segment.");
            return; // Exit if parent not found
        }

        const segmentSummaryDiv = segmentBlock.querySelector('.segment-item-summary');
        if (!segmentSummaryDiv) {
             console.error("Could not find segment summary div.");
             return; // Exit if summary div not found
        }

        // Read parent item options
        const type = itemBlock.querySelector('.item-structure-type').value;
        const itemNumber = itemBlock.dataset.itemId.split('-')[1]; // Get item number from item ID like 'item-1' -> '1'

        let summaryText = `${getItemTypeDescription(type)} #${itemNumber} - `;
        if (type === 'muro') {
            const facesInput = itemBlock.querySelector('.item-faces');
            const faces = facesInput && !facesInput.closest('.hidden') ? parseInt(facesInput.value) : 1;

            const cara1PanelSelect = itemBlock.querySelector('.item-cara1-panel-type');
            const cara1PanelType = cara1PanelSelect && !cara1PanelSelect.closest('.hidden') ? cara1PanelSelect.value : 'N/A';

            const cara2PanelSelect = itemBlock.querySelector('.item-cara2-panel-type');
            const cara2PanelType = (faces === 2 && cara2PanelSelect && !cara2PanelSelect.closest('.hidden')) ? cara2PanelSelect.value : 'N/A';

            const postSpacingInput = itemBlock.querySelector('.item-post-spacing');
            const postSpacing = postSpacingInput && !postSpacingInput.closest('.hidden') ? parseFloat(postSpacingInput.value) : NaN;

            const isDoubleStructureInput = itemBlock.querySelector('.item-double-structure');
            const isDoubleStructure = isDoubleStructureInput && !isDoubleStructureInput.closest('.hidden') ? isDoubleStructureInput.checked : false;

            summaryText += `${faces} Cara${faces > 1 ? 's' : ''}, Panel C1: ${cara1PanelType}`;
            if (faces === 2) summaryText += `, Panel C2: ${cara2PanelType}`;
            if (!isNaN(postSpacing)) summaryText += `, Esp: ${postSpacing.toFixed(2)}m`;
            if (isDoubleStructure) summaryText += `, Doble Estructura`;

         } else if (type === 'cielo') {
            const cieloPanelSelect = itemBlock.querySelector('.item-cielo-panel-type');
            const cieloPanelType = cieloPanelSelect && !cieloPanelSelect.closest('.hidden') ? cieloPanelSelect.value : 'N/A';

            const plenumInput = itemBlock.querySelector('.item-plenum');
            const plenum = plenumInput && !plenumInput.closest('.hidden') ? parseFloat(plenumInput.value) : NaN;

            const angularDeductionInput = itemBlock.querySelector('.item-angular-deduction');
            const angularDeduction = angularDeductionInput && !angularDeductionInput.closest('.hidden') ? parseFloat(angularDeductionInput.value) : NaN;

            summaryText += `Panel: ${cieloPanelType}`;
            if (!isNaN(plenum)) summaryText += `, Pleno: ${plenum.toFixed(2)}m`;
            if (!isNaN(angularDeduction) && angularDeduction > 0) summaryText += `, Desc. Ang: ${angularDeduction.toFixed(2)}m`;

         } else if (type === 'cenefa') { // New Cenefa summary details
             const orientationSelect = itemBlock.querySelector('.item-cenefa-orientation');
             const orientation = orientationSelect && !orientationSelect.closest('.hidden') ? orientationSelect.value : 'N/A';

             const panelTypeSelect = itemBlock.querySelector('.item-cenefa-panel-type');
             const panelType = panelTypeSelect && !panelTypeSelect.closest('.hidden') ? panelTypeSelect.value : 'N/A';

             const sidesInput = itemBlock.querySelector('.item-cenefa-sides');
             const sides = sidesInput && !sidesInput.closest('.hidden') ? parseInt(sidesInput.value) : NaN;

             const plenumInput = itemBlock.querySelector('.item-plenum'); // Re-use plenum input for cenefa if applicable
             const plenum = plenumInput && !plenumInput.closest('.hidden') ? parseFloat(plenumInput.value) : NaN;

             const listonSpacingInput = itemBlock.querySelector('.item-cenefa-liston-spacing');
             const listonSpacing = listonSpacingInput && !listonSpacingInput.closest('.hidden') ? parseFloat(listonSpacingInput.value) : NaN;


             summaryText += `Orientación: ${orientation}, Panel: ${panelType}`;
             if (!isNaN(sides) && sides > 0) summaryText += `, Lados: ${sides}`;
             if (!isNaN(plenum)) summaryText += `, Pleno: ${plenum.toFixed(2)}m`;
             if (!isNaN(listonSpacing) && listonSpacing > 0) summaryText += `, Esp Listón: ${listonSpacing.toFixed(2)}m`;


        } else {
             summaryText += "Configuración Desconocida"; // Fallback for unknown type
        }

        // Update the text content of the dedicated summary div within the segment
        segmentSummaryDiv.textContent = summaryText;
     };

    // --- Function to Import Segments from Excel ---
    const importSegmentsFromExcel = (itemBlock, file, itemType) => {
        console.log(`Importing segments from Excel for ${itemType} item...`);
        // Check if xlsx library is loaded (assumed to be included via a <script> tag in index.html)
        if (typeof XLSX === 'undefined') {
             alert("Error al importar: La librería xlsx no está cargada correctamente.");
             console.error("XLSX library is not loaded.");
             return;
        }

        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                // Assuming the first sheet is the one with dimensions
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                // Convert sheet to array of arrays, starting from row 1 (index 0)
                // `header: 1` means treat the first row as data, not header names for object keys.
                const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                if (jsonSheet.length <= 1) { // Need at least a header row and one data row
                     alert("El archivo Excel está vacío o solo contiene encabezados.");
                     return;
                }

                // Assuming the first row is the header (index 0)
                const headerRow = jsonSheet[0];
                let widthColIndex = -1;
                let heightLengthColIndex = -1; // Will be 'Alto' for muro, 'Largo' for cielo, 'Alto' or 'Ancho' for cenefa depending on context (need clarification from user/matriz)
                let lengthColIndex = -1; // For Cenefa, might need Largo, Ancho, Alto

                // Determine required headers based on item type
                let requiredHeaders = ['ancho'];
                if (itemType === 'muro') requiredHeaders.push('alto');
                else if (itemType === 'cielo') requiredHeaders.push('largo');
                else if (itemType === 'cenefa') {
                     // For cenefa, the dimensions could be Largo, Ancho, Alto.
                     // Let's assume the most common case: importing Largo, Ancho, Alto columns.
                     requiredHeaders.push('largo', 'ancho', 'alto');
                     // We'll use 'ancho' and 'largo'/'alto' from the input fields, so we need to map imported columns to these.
                     // Let's adjust column mapping based on Cenefa inputs ( Largo, Ancho, Alto)
                     lengthColIndex = -1; // Use 'length' for Cenefa Largo
                     widthColIndex = -1; // Use 'width' for Cenefa Ancho
                     heightLengthColIndex = -1; // Use 'height' for Cenefa Alto
                     requiredHeaders = ['largo', 'ancho', 'alto']; // Update required headers for Cenefa
                }


                // Find column indices (case-insensitive, trimmed)
                headerRow.forEach((header, index) => {
                     if (typeof header === 'string') { // Ensure header is a string
                         const cleanHeader = header.trim().toLowerCase();
                          if (itemType === 'cenefa') {
                               if (cleanHeader === 'largo') lengthColIndex = index;
                               else if (cleanHeader === 'ancho') widthColIndex = index;
                               else if (cleanHeader === 'alto') heightLengthColIndex = index; // Using heightLengthColIndex for 'alto' in cenefa
                          } else { // muro or cielo
                              if (cleanHeader === 'ancho') widthColIndex = index;
                              else if (cleanHeader === requiredHeaders[1]) heightLengthColIndex = index; // 'alto' or 'largo'
                          }
                     }
                });

                // Validate header presence based on item type
                let headersValid = false;
                let missingHeaders = [];
                 if (itemType === 'muro' || itemType === 'cielo') {
                      if (widthColIndex !== -1 && heightLengthColIndex !== -1) {
                           headersValid = true;
                       } else {
                           if (widthColIndex === -1) missingHeaders.push('"Ancho"');
                           if (heightLengthColIndex === -1) missingHeaders.push(`"${requiredHeaders[1].charAt(0).toUpperCase() + requiredHeaders[1].slice(1)}"`); // Capitalize first letter
                       }
                 } else if (itemType === 'cenefa') {
                     // For Cenefa, we need Largo, Ancho, and Alto.
                     if (lengthColIndex !== -1 && widthColIndex !== -1 && heightLengthColIndex !== -1) {
                          headersValid = true;
                     } else {
                          if (lengthColIndex === -1) missingHeaders.push('"Largo"');
                          if (widthColIndex === -1) missingHeaders.push('"Ancho"');
                          if (heightLengthColIndex === -1) missingHeaders.push('"Alto"');
                     }
                 }


                if (!headersValid) {
                    alert(`Error: El archivo Excel debe contener las columnas ${missingHeaders.join(', ')} en la primera fila para ${getItemTypeName(itemType)}.`);
                    return;
                }


                 // Get the correct segments list container based on item type
                const segmentsListContainer = itemBlock.querySelector(itemType === 'muro' ? '.muro-segments .segments-list' : (itemType === 'cielo' ? '.cielo-segments .segments-list' : '.cenefa-segments .segments-list')); // Added .cenefa-segments
                const createSegmentFunc = itemType === 'muro' ? createMuroSegmentBlock : (itemType === 'cielo' ? createCieloSegmentBlock : createCenefaSegmentBlock); // Added createCenefaSegmentBlock

                // --- Opcional: Preguntar si desea reemplazar segmentos existentes ---
                // const replaceExisting = confirm("¿Desea reemplazar los segmentos existentes con los datos del Excel?");
                // if (replaceExisting) {
                //      segmentsListContainer.innerHTML = '';
                //      // Reset segment numbers later during adding
                // }
                 // --- Fin Opcional ---


                let segmentsImportedCount = 0;
                let invalidRows = [];
                let currentSegmentCount = segmentsListContainer.querySelectorAll(itemType === 'muro' ? '.muro-segment' : (itemType === 'cielo' ? '.cielo-segment' : '.cenefa-segment')).length; // Added .cenefa-segment

                // Process data rows (start from index 1, skipping header row at index 0)
                for (let i = 1; i < jsonSheet.length; i++) {
                    const row = jsonSheet[i];
                    // Ensure row is an array and has enough columns to cover the required indices
                    let requiredIndex = -1;
                     if (itemType === 'cenefa') {
                          requiredIndex = Math.max(lengthColIndex, widthColIndex, heightLengthColIndex);
                     } else { // muro or cielo
                          requiredIndex = Math.max(widthColIndex, heightLengthColIndex);
                     }

                    if (!Array.isArray(row) || row.length <= requiredIndex) {
                         invalidRows.push(`Fila ${i + 1}: Datos incompletos.`);
                         continue; // Skip to next row
                    }

                    // Read values - use || 0 to handle potential undefined/null if column exists but cell is empty
                    // Validate and parse dimensions based on item type
                    let dim1, dim2, dim3; // General variables for dimensions

                    if (itemType === 'muro') {
                        dim1 = parseFloat(row[widthColIndex]); // Ancho
                        dim2 = parseFloat(row[heightLengthColIndex]); // Alto
                         if (isNaN(dim1) || dim1 <= 0 || isNaN(dim2) || dim2 <= 0) {
                              invalidRows.push(`Fila ${i + 1}: Dimensiones de Muro inválidas (Ancho y Alto deben ser > 0)`);
                              continue;
                         }
                    } else if (itemType === 'cielo') {
                        dim1 = parseFloat(row[widthColIndex]); // Ancho
                        dim2 = parseFloat(row[heightLengthColIndex]); // Largo
                         if (isNaN(dim1) || dim1 <= 0 || isNaN(dim2) || dim2 <= 0) {
                              invalidRows.push(`Fila ${i + 1}: Dimensiones de Cielo inválidas (Ancho y Largo deben ser > 0)`);
                              continue;
                         }
                    } else if (itemType === 'cenefa') {
                        dim1 = parseFloat(row[lengthColIndex]); // Largo
                        dim2 = parseFloat(row[widthColIndex]); // Ancho
                        dim3 = parseFloat(row[heightLengthColIndex]); // Alto
                        // For Cenefa, we need Largo > 0, Ancho > 0, Alto > 0
                         if (isNaN(dim1) || dim1 <= 0 || isNaN(dim2) || dim2 <= 0 || isNaN(dim3) || dim3 <= 0) {
                             invalidRows.push(`Fila ${i + 1}: Dimensiones de Cenefa inválidas (Largo, Ancho y Alto deben ser > 0).`);
                             continue;
                         }
                    } else {
                         invalidRows.push(`Fila ${i + 1}: Tipo de ítem desconocido.`);
                         continue;
                    }


                    // Create and add the new segment
                    currentSegmentCount++; // Increment segment count for the new segment number
                    const newSegment = createSegmentFunc(itemBlock.dataset.itemId, currentSegmentCount); // Pass item ID and new segment number

                    // Populate dimensions based on item type
                    if (itemType === 'muro') {
                         newSegment.querySelector('.item-width').value = dim1.toFixed(2);
                         newSegment.querySelector('.item-height').value = dim2.toFixed(2);
                    } else if (itemType === 'cielo') {
                         newSegment.querySelector('.item-width').value = dim1.toFixed(2);
                         newSegment.querySelector('.item-length').value = dim2.toFixed(2);
                    } else if (itemType === 'cenefa') {
                         newSegment.querySelector('.item-length').value = dim1.toFixed(2); // Cenefa Largo
                         newSegment.querySelector('.item-width').value = dim2.toFixed(2); // Cenefa Ancho
                         newSegment.querySelector('.item-height').value = dim3.toFixed(2); // Cenefa Alto
                    }


                    segmentsListContainer.appendChild(newSegment);
                    updateSegmentItemSummary(newSegment); // Update summary for the new segment

                    segmentsImportedCount++;
                }

                // Re-number all segments visually after import (needed if replacing or adding)
                const segmentSelector = itemType === 'muro' ? '.muro-segment' : (itemType === 'cielo' ? '.cielo-segment' : '.cenefa-segment'); // Added .cenefa-segment
                segmentsListContainer.querySelectorAll(`${segmentSelector} h4`).forEach((h4, index) => {
                    h4.textContent = `Segmento ${index + 1}`;
                 });

                // Provide feedback to the user
                let feedbackMessage = `Importación completada. Se agregaron ${segmentsImportedCount} segmento(s) al ${getItemTypeName(itemType)} #${itemBlock.dataset.itemId.split('-')[1]}.`;
                if (invalidRows.length > 0) {
                     feedbackMessage += `\nHubo ${invalidRows.length} fila(s) con errores (no importadas). Revisa la consola del navegador para más detalles.`;
                     alert(feedbackMessage); // Use alert for errors or partial import
                     console.warn("Filas no importadas:", invalidRows);
                } else {
                     alert(feedbackMessage); // Use alert for successful import
                }


                // Clear previous calculation results as dimensions have changed
                resultsContent.innerHTML = '<p>Dimensiones importadas. Recalcula los materiales totales.</p>';
                downloadOptionsDiv.classList.add('hidden');
                lastCalculatedTotalMaterials = {};
                lastCalculatedItemsSpecs = [];
                lastErrorMessages = [];
                lastCalculatedWorkArea = ''; // Clear stored data on import

            } catch (error) {
                console.error("Error reading or processing Excel file:", error);
                alert(`Error al leer o procesar el archivo Excel: ${error.message}`);
            }
        };
        reader.onerror = (error) => {
            console.error("Error reading file:", error);
            alert("Error al leer el archivo.");
        };

        // Read the file as an ArrayBuffer
        reader.readAsArrayBuffer(file);
    };


     // --- Function to Create a Muro Segment Input Block ---
    const createMuroSegmentBlock = (itemId, segmentNumber) => {
        // HTML structure for a wall segment. Uses classes defined in style.css (.muro-segment, .segment-header-line, .segment-item-summary, .input-group, .remove-segment-btn)
        const segmentHtml = `
            <div class="muro-segment" data-segment-id="${itemId}-mseg-${segmentNumber}">
                 <div class="segment-header-line"> <h4>Segmento ${segmentNumber}</h4>
                    <div class="segment-item-summary"></div> </div>
                 <button type="button" class="remove-segment-btn">X</button>
                 <div class="input-group">
                    <label for="mwidth-${itemId}-mseg-${segmentNumber}">Ancho (m):</label>
                    <input type="number" class="item-width" id="mwidth-${itemId}-mseg-${segmentNumber}" step="0.01" min="0" value="3.0">
                </div>
                <div class="input-group">
                    <label for="mheight-${itemId}-mseg-${segmentNumber}">Alto (m):</label>
                    <input type="number" class="item-height" id="mheight-${itemId}-mseg-${segmentNumber}" step="0.01" min="0" value="2.4">
                </div>
            </div>
        `;
        const newElement = document.createElement('div');
        newElement.innerHTML = segmentHtml.trim();
        const segmentBlock = newElement.firstChild; // Get the actual div element

        // Add remove listener
        const removeButton = segmentBlock.querySelector('.remove-segment-btn');
        removeButton.addEventListener('click', () => {
            const segmentsContainer = segmentBlock.closest('.segments-list'); // Correct selector
            const itemBlock = segmentsContainer.closest('.item-block'); // Get parent item for re-numbering
            const segmentTypeSelector = itemBlock.querySelector('.item-structure-type').value === 'muro' ? '.muro-segment' : '.cielo-segment'; // Updated for cenefa later


            if (segmentsContainer.querySelectorAll(segmentTypeSelector).length > 1) {
                 segmentBlock.remove();

                 // Re-number segments visually after removal
                 segmentsContainer.querySelectorAll(`${segmentTypeSelector} h4`).forEach((h4, index) => {
                    h4.textContent = `Segmento ${index + 1}`;
                 });

                 // Clear results and hide download buttons after removal
                 resultsContent.innerHTML = '<p>Segmento eliminado. Recalcula los materiales totales.</p>';
                 downloadOptionsDiv.classList.add('hidden');
                 lastCalculatedTotalMaterials = {};
                 lastCalculatedItemsSpecs = [];
                 lastErrorMessages = [];
                 lastCalculatedWorkArea = ''; // Clear stored data on item removal
            } else {
                 alert(`Un ${getItemTypeName(itemBlock.querySelector('.item-structure-type').value)} debe tener al menos un segmento.`);
            }
         });

        return segmentBlock;
    };

     // --- Function to Create a Cielo Segment Input Block ---
     const createCieloSegmentBlock = (itemId, segmentNumber) => {
         // HTML structure for a ceiling segment. Uses classes defined in style.css (.cielo-segment, .segment-header-line, .segment-item-summary, .input-group, .remove-segment-btn)
         const segmentHtml = `
            <div class="cielo-segment" data-segment-id="${itemId}-cseg-${segmentNumber}">
                 <div class="segment-header-line"> <h4>Segmento ${segmentNumber}</h4>
                    <div class="segment-item-summary"></div> </div>
                 <button type="button" class="remove-segment-btn">X</button>
                 <div class="input-group">
                    <label for="cwidth-${itemId}-cseg-${segmentNumber}">Ancho (m):</label>
                    <input type="number" class="item-width" id="cwidth-${itemId}-cseg-${segmentNumber}" step="0.01" min="0" value="3.0">
                </div>
                <div class="input-group">
                    <label for="clength-${itemId}-cseg-${segmentNumber}">Largo (m):</label>
                    <input type="number" class="item-length" id="clength-${itemId}-cseg-${segmentNumber}" step="0.01" min="0" value="4.0">
                </div>
            </div>
        `;
        const newElement = document.createElement('div');
        newElement.innerHTML = segmentHtml.trim();
        const segmentBlock = newElement.firstChild; // Get the actual div element

         // Add remove listener
         const removeButton = segmentBlock.querySelector('.remove-segment-btn');
         removeButton.addEventListener('click', () => {
             const segmentsContainer = segmentBlock.closest('.segments-list'); // Correct selector
             const itemBlock = segmentsContainer.closest('.item-block'); // Get parent item for re-numbering
             const segmentTypeSelector = itemBlock.querySelector('.item-structure-type').value === 'muro' ? '.muro-segment' : '.cielo-segment'; // Updated for cenefa later


             if (segmentsContainer.querySelectorAll(segmentTypeSelector).length > 1) {
                  segmentBlock.remove();
                  // Re-number segments visually after removal
                  segmentsContainer.querySelectorAll(`${segmentTypeSelector} h4`).forEach((h4, index) => {
                     h4.textContent = `Segmento ${index + 1}`;
                  });

                  // Clear results and hide download buttons after removal
                  resultsContent.innerHTML = '<p>Segmento eliminado. Recalcula los materiales totales.</p>';
                  downloadOptionsDiv.classList.add('hidden');
                  lastCalculatedTotalMaterials = {};
                  lastCalculatedItemsSpecs = [];
                  lastErrorMessages = [];
                  lastCalculatedWorkArea = ''; // Clear stored data on item removal
             } else {
                  alert(`Un ${getItemTypeName(itemBlock.querySelector('.item-structure-type').value)} debe tener al menos un segmento.`);
             }
         });

         return segmentBlock;
     };

     // --- Function to Create a Cenefa Segment Input Block ---
     // New function for Cenefa segments
     const createCenefaSegmentBlock = (itemId, segmentNumber) => {
          // HTML structure for a cenefa segment. Requires classes like .cenefa-segment, .segment-header-line, .segment-item-summary, .input-group, .remove-segment-btn
         const segmentHtml = `
            <div class="cenefa-segment" data-segment-id="${itemId}-cseg-${segmentNumber}">
                 <div class="segment-header-line"> <h4>Segmento ${segmentNumber}</h4>
                    <div class="segment-item-summary"></div> </div>
                 <button type="button" class="remove-segment-btn">X</button>
                 <div class="input-group">
                    <label for="cenefa-largo-${itemId}-cseg-${segmentNumber}">Largo (m):</label>
                    <input type="number" class="item-length" id="cenefa-largo-${itemId}-cseg-${segmentNumber}" step="0.01" min="0" value="2.0">
                </div>
                 <div class="input-group">
                    <label for="cenefa-ancho-${itemId}-cseg-${segmentNumber}">Ancho (m):</label>
                    <input type="number" class="item-width" id="cenefa-ancho-${itemId}-cseg-${segmentNumber}" step="0.01" min="0" value="0.30">
                </div>
                 <div class="input-group">
                    <label for="cenefa-alto-${itemId}-cseg-${segmentNumber}">Alto (m):</label>
                    <input type="number" class="item-height" id="cenefa-alto-${itemId}-cseg-${segmentNumber}" step="0.01" min="0" value="0.40">
                </div>
            </div>
        `;
        const newElement = document.createElement('div');
        newElement.innerHTML = segmentHtml.trim();
        const segmentBlock = newElement.firstChild;

         // Add remove listener (similar to muro/cielo segments)
         const removeButton = segmentBlock.querySelector('.remove-segment-btn');
         removeButton.addEventListener('click', () => {
             const segmentsContainer = segmentBlock.closest('.segments-list');
             const itemBlock = segmentsContainer.closest('.item-block');
             const segmentTypeSelector = '.cenefa-segment'; // Specific selector for cenefa segments

             if (segmentsContainer.querySelectorAll(segmentTypeSelector).length > 1) {
                  segmentBlock.remove();
                  // Re-number segments visually after removal
                  segmentsContainer.querySelectorAll(`${segmentTypeSelector} h4`).forEach((h4, index) => {
                     h4.textContent = `Segmento ${index + 1}`;
                  });

                  // Clear results and hide download buttons after removal
                  resultsContent.innerHTML = '<p>Segmento de cenefa eliminado. Recalcula los materiales totales.</p>';
                  downloadOptionsDiv.classList.add('hidden');
                  lastCalculatedTotalMaterials = {};
                  lastCalculatedItemsSpecs = [];
                  lastErrorMessages = [];
                  lastCalculatedWorkArea = ''; // Clear stored data on item removal
             } else {
                  alert(`Una Cenefa debe tener al menos un segmento.`);
             }
         });

         return segmentBlock;
     };


    // --- Function to Update Input Visibility WITHIN an Item Block ---
    // Uses the 'hidden' class defined in style.css
    const updateItemInputVisibility = (itemBlock) => {
        const structureTypeSelect = itemBlock.querySelector('.item-structure-type');

        // Common input groups (some are type-specific and might be hidden)
        const facesInputGroup = itemBlock.querySelector('.item-faces-input');
        const muroPanelTypesDiv = itemBlock.querySelector('.muro-panel-types');
        const cieloPanelTypeDiv = itemBlock.querySelector('.cielo-panel-type');
        const cenefaPanelTypeDiv = itemBlock.querySelector('.cenefa-panel-type'); // New: Cenefa panel type div
        const cenefaOrientationInputGroup = itemBlock.querySelector('.item-cenefa-orientation-input'); // New: Cenefa orientation input group
        const cenefaSidesInputGroup = itemBlock.querySelector('.item-cenefaSides-input'); // New: Cenefa sides input group
        const postSpacingInputGroup = itemBlock.querySelector('.item-post-spacing-input');
        const doubleStructureInputGroup = itemBlock.querySelector('.item-double-structure-input');
        const plenumInputGroup = itemBlock.querySelector('.item-plenum-input'); // Used for Cielo and Cenefa
        const angularDeductionInputGroup = itemBlock.querySelector('.item-angular-deduction-input'); // Used for Cielo
        // Cenefa might have its own spacing inputs if needed based on detailed structure
        // Let's add inputs for Cenefa spacing if they are defined in the HTML structure
        const cenefaListonSpacingInputGroup = itemBlock.querySelector('.item-cenefa-liston-spacing-input'); // New


        // Type-specific dimension/segment containers
        const muroSegmentsContainer = itemBlock.querySelector('.muro-segments');
        const cieloSegmentsContainer = itemBlock.querySelector('.cielo-segments');
        const cenefaSegmentsContainer = itemBlock.querySelector('.cenefa-segments'); // New


        // --- Referencias a los contenedores de importación ---
        const importMuroContainer = itemBlock.querySelector('.import-excel-segments-muro');
        const importCieloContainer = itemBlock.querySelector('.import-excel-segments-cielo');
        const importCenefaContainer = itemBlock.querySelector('.import-excel-segments-cenefa'); // New
        // --- Fin Referencias ---


        const type = structureTypeSelect.value;

        // Reset visibility for ALL type-specific input groups within this block by adding 'hidden' class
        facesInputGroup.classList.add('hidden');
        muroPanelTypesDiv.classList.add('hidden');
        cieloPanelTypeDiv.classList.add('hidden');
        if (cenefaPanelTypeDiv) cenefaPanelTypeDiv.classList.add('hidden'); // Hide new cenefa div
        if (cenefaOrientationInputGroup) cenefaOrientationInputGroup.classList.add('hidden'); // Hide new cenefa orientation
        if (cenefaSidesInputGroup) cenefaSidesInputGroup.classList.add('hidden'); // Hide new cenefa sides
        postSpacingInputGroup.classList.add('hidden');
        doubleStructureInputGroup.classList.add('hidden');
        plenumInputGroup.classList.add('hidden'); // Plenum is hidden by default, shown for cielo/cenefa
        muroSegmentsContainer.classList.add('hidden');
        cieloSegmentsContainer.classList.add('hidden');
        if (cenefaSegmentsContainer) cenefaSegmentsContainer.classList.add('hidden'); // Hide new cenefa segments
        angularDeductionInputGroup.classList.add('hidden'); // Angular deduction is only for Cielo
        if (cenefaListonSpacingInputGroup) cenefaListonSpacingInputGroup.classList.add('hidden'); // Hide new cenefa spacing


        // --- Oculta todos los botones de importación inicialmente ---
        if (importMuroContainer) importMuroContainer.classList.add('hidden');
        if (importCieloContainer) importCieloContainer.classList.add('hidden');
        if (importCenefaContainer) importCenefaContainer.classList.add('hidden'); // Hide new cenefa import
        // --- Fin Ocultar ---


        // Set visibility based on selected type for THIS block by removing 'hidden' class
        if (type === 'muro') {
            facesInputGroup.classList.remove('hidden');
            muroPanelTypesDiv.classList.remove('hidden'); // Show wall panel type selectors
            postSpacingInputGroup.classList.remove('hidden'); // Post spacing applies to walls
            doubleStructureInputGroup.classList.remove('hidden'); // Double structure applies to walls
            muroSegmentsContainer.classList.remove('hidden'); // Show muro segments container
             // --- Muestra el botón de importación para Muro ---
             if (importMuroContainer) importMuroContainer.classList.remove('hidden');
             // --- Fin Mostrar ---

             // Update visibility of face-specific panel type selectors based on faces input
             const facesInput = itemBlock.querySelector('.item-faces');
             const cara2PanelTypeGroup = itemBlock.querySelector('.cara-2-panel-type-group');

             if (parseInt(facesInput.value) === 2) {
                 cara2PanelTypeGroup.classList.remove('hidden');
             } else {
                 cara2PanelTypeGroup.classList.add('hidden');
             }

             // --- Llama a la función para actualizar el resumen dentro de cada segmento de muro ---
             // Esto se hace para que los resúmenes se actualicen si cambias el tipo de item a muro
             itemBlock.querySelectorAll('.muro-segment').forEach(segBlock => {
                 updateSegmentItemSummary(segBlock);
             });
             // --- Fin llamado ---


        } else if (type === 'cielo') {
            cieloPanelTypeDiv.classList.remove('hidden'); // Show ceiling panel type selector
            plenumInputGroup.classList.remove('hidden');
            cieloSegmentsContainer.classList.remove('hidden'); // Show cielo segments container
            angularDeductionInputGroup.classList.remove('hidden'); // Muestra el input de descuento angular para cielos
            // --- Muestra el botón de importación para Cielo ---
             if (importCieloContainer) importCieloContainer.classList.remove('hidden');
             // --- Fin Mostrar ---

            // --- Llama a la función para actualizar el resumen dentro de cada segmento de cielo ---
             // Esto se hace para que los resúmenes se actualicen si cambias el tipo a 'cielo'.
            itemBlock.querySelectorAll('.cielo-segment').forEach(segBlock => {
                updateSegmentItemSummary(segBlock);
            });
            // --- Fin llamado ---

        } else if (type === 'cenefa') { // Show inputs for Cenefa
             if (cenefaPanelTypeDiv) cenefaPanelTypeDiv.classList.remove('hidden');
             if (cenefaOrientationInputGroup) cenefaOrientationInputGroup.classList.remove('hidden');
             if (cenefaSidesInputGroup) cenefaSidesInputGroup.classList.remove('hidden');
             if (cenefaSegmentsContainer) cenefaSegmentsContainer.classList.remove('hidden');
             if (cenefaListonSpacingInputGroup) cenefaListonSpacingInputGroup.classList.remove('hidden'); // Show cenefa spacing
              // --- Muestra el botón de importación para Cenefa ---
             if (importCenefaContainer) importCenefaContainer.classList.remove('hidden');
             // --- Fin Mostrar ---

             // --- Llama a la función para actualizar el resumen dentro de cada segmento de cenefa ---
             itemBlock.querySelectorAll('.cenefa-segment').forEach(segBlock => {
                 updateSegmentItemSummary(segBlock);
             });
             // --- Fin llamado ---
        }
        // No need for 'else' as all are hidden by default initially
    };

    // --- Function to update the main item header summary (kept for consistency) ---
    // The detailed summary is handled within each segment block.
    const updateItemHeaderSummary = (itemBlock) => {
         const itemHeader = itemBlock.querySelector('h3');
         const type = itemBlock.querySelector('.item-structure-type').value;
         const itemNumber = itemBlock.dataset.itemId.split('-')[1]; // Get item number from item ID

         // Update the main header to just show Type and Number
         itemHeader.textContent = `${getItemTypeDescription(type)} #${itemNumber}`;
     };


    // --- Function to Create an Item Input Block ---
    // Creates the main container for an item (wall or ceiling or cenefa) using classes from style.css (.item-block, .remove-item-btn, .input-group, .hidden)
    const createItemBlock = () => {
        itemCounter++;
        const itemId = `item-${itemCounter}`;

        // Restructured HTML template for an item block
        const itemHtml = `
            <div class="item-block" data-item-id="${itemId}">
                <h3>${getItemTypeDescription('muro')} #${itemCounter}</h3> <button class="remove-item-btn">Eliminar</button>

                <div class="input-group">
                    <label for="type-${itemId}">Tipo de Estructura:</label>
                    <select class="item-structure-type" id="type-${itemId}">
                        <option value="muro">Muro</option>
                        <option value="cielo">Cielo Falso</option>
                        <option value="cenefa">Cenefa</option> </select>
                </div>

                <div class="input-group item-faces-input">
                    <label for="faces-${itemId}">Nº de Caras (1 o 2):</label>
                    <input type="number" class="item-faces" id="faces-${itemId}" step="1" min="1" max="2" value="1">
                </div>

                <div class="muro-panel-types">
                    <div class="input-group cara-1-panel-type-group">
                        <label for="cara1-panel-type-${itemId}">Panel Cara 1:</label>
                        <select class="item-cara1-panel-type" id="cara1-panel-type-${itemId}"></select>
                    </div>
                    <div class="input-group cara-2-panel-type-group hidden">
                        <label for="cara2-panel-type-${itemId}">Panel Cara 2:</label>
                        <select class="item-cara2-panel-type" id="cara2-panel-type-${itemId}"></select>
                    </div>
                </div>

                <div class="input-group item-post-spacing-input">
                    <label for="post-spacing-${itemId}">Espaciamiento Postes (m):</label>
                    <input type="number" class="item-post-spacing" id="post-spacing-${itemId}" step="0.01" min="0.1" value="0.40">
                </div>

                <div class="input-group item-double-structure-input">
                    <label for="double-structure-${itemId}">Estructura Doble:</label>
                    <input type="checkbox" class="item-double-structure" id="double-structure-${itemId}">
                </div>


                 <div class="input-group cielo-panel-type hidden">
                     <label for="cielo-panel-type-${itemId}">Tipo de Panel:</label>
                    <select class="item-cielo-panel-type" id="cielo-panel-type-${itemId}"></select>
                </div>

                <div class="input-group item-plenum-input hidden">
                    <label for="plenum-${itemId}">Pleno (m):</label> <input type="number" class="item-plenum" id="plenum-${itemId}" step="0.01" min="0" value="0.5">
                </div>

                <div class="input-group item-angular-deduction-input hidden">
                    <label for="angular-deduction-${itemId}">Metros a descontar de Angular:</label>
                    <input type="number" class="item-angular-deduction" id="angular-deduction-${itemId}" step="0.01" min="0" value="0">
                </div>

                <div class="input-group item-cenefa-orientation-input hidden">
                    <label for="cenefa-orientation-${itemId}">Orientación Principal:</label>
                    <select class="item-cenefa-orientation" id="cenefa-orientation-${itemId}">
                         <option value="horizontal">Horizontal</option>
                         <option value="vertical">Vertical</option>
                    </select>
                </div>

                <div class="input-group cenefa-panel-type hidden">
                     <label for="cenefa-panel-type-${itemId}">Tipo de Panel:</label>
                    <select class="item-cenefa-panel-type" id="cenefa-panel-type-${itemId}"></select>
                </div>

                 <div class="input-group item-cenefaSides-input hidden">
                    <label for="cenefa-sides-${itemId}">Nº de Lados/Caras con Panel:</label>
                    <input type="number" class="item-cenefa-sides" id="cenefa-sides-${itemId}" step="1" min="1" value="1">
                </div>

                 <div class="input-group item-cenefa-liston-spacing-input hidden">
                    <label for="cenefa-liston-spacing-${itemId}">Espaciamiento Canal Listón (m):</label>
                    <input type="number" class="item-cenefa-liston-spacing" id="cenefa-liston-spacing-${itemId}" step="0.01" min="0.1" value="0.40">
                </div>

                <div class="muro-segments">
                    <h4>Segmentos del Muro:</h4>
                     <div class="segments-list">
                     </div>
                    <button type="button" class="add-segment-btn">Agregar Segmento</button>
                     <div class="import-excel-segments-muro">
                        <label for="import-muro-${itemId}" class="button">Importar Ancho/Alto (Excel)</label>
                        <input type="file" id="import-muro-${itemId}" class="import-segments-input hidden" accept=".xlsx">
                    </div>
                    </div>

                 <div class="cielo-segments hidden">
                    <h4>Segmentos del Cielo Falso:</h4>
                     <div class="segments-list">
                         </div>
                    <button type="button" class="add-segment-btn">Agregar Segmento</button>
                     <div class="import-excel-segments-cielo">
                        <label for="import-cielo-${itemId}" class="button">Importar Ancho/Largo (Excel)</label>
                        <input type="file" id="import-cielo-${itemId}" class="import-segments-input hidden" accept=".xlsx">
                    </div>
                    </div>

                <div class="cenefa-segments hidden">
                    <h4>Segmentos de Cenefa:</h4>
                     <div class="segments-list">
                         </div>
                    <button type="button" class="add-segment-btn">Agregar Segmento</button>
                     <div class="import-excel-segments-cenefa">
                        <label for="import-cenefa-${itemId}" class="button">Importar Largo/Ancho/Alto (Excel)</label>
                        <input type="file" id="import-cenefa-${itemId}" class="import-segments-input hidden" accept=".xlsx">
                    </div>
                    </div>
                </div>
        `;

        const newElement = document.createElement('div');
        newElement.innerHTML = itemHtml.trim();
        const itemBlock = newElement.firstChild; // Get the actual div element

        itemsContainer.appendChild(itemBlock);

        // Actualiza el encabezado principal del ítem al crearlo
        updateItemHeaderSummary(itemBlock);

        // Add an initial segment based on the DEFAULT type ('muro')
        const muroSegmentsListContainer = itemBlock.querySelector('.muro-segments .segments-list');
        if (muroSegmentsListContainer) {
             const initialSegment = createMuroSegmentBlock(itemId, 1);
             muroSegmentsListContainer.appendChild(initialSegment);
             // --- Llama a la función de resumen para el segmento inicial de muro ---
             // Esto se hace para que el resumen aparezca al crear el item por defecto
             updateSegmentItemSummary(initialSegment);
             // --- Fin llamado ---


             // Add listener for "Agregar Segmento" button for muro
             const addMuroSegmentBtn = itemBlock.querySelector('.muro-segments .add-segment-btn');
             addMuroSegmentBtn.addEventListener('click', () => {
                 const currentSegments = muroSegmentsListContainer.querySelectorAll('.muro-segment').length;
                 const newSegment = createMuroSegmentBlock(itemId, currentSegments + 1);
                 muroSegmentsListContainer.appendChild(newSegment);
                 // --- Llama a la función de resumen para el nuevo segmento de muro ---
                 updateSegmentItemSummary(newSegment);
                 // --- Fin llamado ---

                 // Clear results and hide download buttons after adding a segment
                 resultsContent.innerHTML = '<p>Segmento de muro agregado. Recalcula los materiales totales.</p>';
                 downloadOptionsDiv.classList.add('hidden');
                 lastCalculatedTotalMaterials = {};
                 lastCalculatedItemsSpecs = [];
                 lastErrorMessages = [];
                 lastCalculatedWorkArea = ''; // Clear stored data on adding/removing segment
             });

             // --- NUEVO: Add event listener for Muro Excel import input ---
             const importMuroInput = itemBlock.querySelector(`#import-muro-${itemId}`);
             if (importMuroInput) {
                 importMuroInput.addEventListener('change', (event) => {
                     const file = event.target.files[0];
                     if (file) {
                         importSegmentsFromExcel(itemBlock, file, 'muro'); // Pass the itemBlock, file, and type
                     }
                     // Reset the input so the same file can be selected again
                     event.target.value = null;
                 });
             }
             // --- FIN NUEVO ---

        }

         // Add listener for "Agregar Segmento" button for cielo (initially hidden)
         const cieloSegmentsListContainer = itemBlock.querySelector('.cielo-segments .segments-list');
         if (cieloSegmentsListContainer) {
              // No creamos un segmento inicial de cielo aquí porque el item por defecto es muro.
             // El segmento inicial de cielo se creará cuando se cambie el tipo a 'cielo'.
             const addCieloSegmentBtn = itemBlock.querySelector('.cielo-segments .add-segment-btn');
             addCieloSegmentBtn.addEventListener('click', () => {
                 const currentSegments = cieloSegmentsListContainer.querySelectorAll('.cielo-segment').length;
                 const newSegment = createCieloSegmentBlock(itemId, currentSegments + 1);
                 cieloSegmentsListContainer.appendChild(newSegment);
                 // --- Llama a la función de resumen para el nuevo segmento de cielo ---
                 updateSegmentItemSummary(newSegment);
                 // --- Fin llamado ---

                 // Clear results and hide download buttons after adding a segment
                 resultsContent.innerHTML = '<p>Segmento de cielo agregado. Recalcula los materiales totales.</p>';
                 downloadOptionsDiv.classList.add('hidden');
                 lastCalculatedTotalMaterials = {};
                 lastCalculatedItemsSpecs = [];
                 lastErrorMessages = [];
                 lastCalculatedWorkArea = ''; // Clear stored data on adding/removing segment
             });

              // --- NUEVO: Add event listener for Cielo Excel import input ---
              const importCieloInput = itemBlock.querySelector(`#import-cielo-${itemId}`);
              if (importCieloInput) {
                   importCieloInput.addEventListener('change', (event) => {
                       const file = event.target.files[0];
                       if (file) {
                           importSegmentsFromExcel(itemBlock, file, 'cielo'); // Pass the itemBlock, file, and type
                       }
                       // Reset the input so the same file can be selected again
                       event.target.value = null;
                   });
              }
              // --- FIN NUEVO ---

          }

         // Add listener for "Agregar Segmento" button for cenefa (initially hidden)
         const cenefaSegmentsListContainer = itemBlock.querySelector('.cenefa-segments .segments-list'); // New
         if (cenefaSegmentsListContainer) {
              // No creamos un segmento inicial de cenefa aquí. Se creará cuando se cambie el tipo.
             const addCenefaSegmentBtn = itemBlock.querySelector('.cenefa-segments .add-segment-btn'); // New
             addCenefaSegmentBtn.addEventListener('click', () => {
                 const currentSegments = cenefaSegmentsListContainer.querySelectorAll('.cenefa-segment').length;
                 const newSegment = createCenefaSegmentBlock(itemId, currentSegments + 1); // New function call
                 cenefaSegmentsListContainer.appendChild(newSegment);
                 updateSegmentItemSummary(newSegment); // Update summary for the new segment

                 // Clear results and hide download buttons after adding a segment
                 resultsContent.innerHTML = '<p>Segmento de cenefa agregado. Recalcula los materiales totales.</p>';
                 downloadOptionsDiv.classList.add('hidden');
                 lastCalculatedTotalMaterials = {};
                 lastCalculatedItemsSpecs = [];
                 lastErrorMessages = [];
                 lastCalculatedWorkArea = ''; // Clear stored data on adding/removing segment
             });

              // --- NUEVO: Add event listener for Cenefa Excel import input ---
              const importCenefaInput = itemBlock.querySelector(`#import-cenefa-${itemId}`); // New
              if (importCenefaInput) {
                   importCenefaInput.addEventListener('change', (event) => {
                       const file = event.target.files[0];
                       if (file) {
                           importSegmentsFromExcel(itemBlock, file, 'cenefa'); // Pass the itemBlock, file, and type
                       }
                       // Reset the input so the same file can be selected again
                       event.target.value = null;
                   });
              }
              // --- FIN NUEVO ---
          }


        // Populate panel type selects in the new block
        const cara1PanelSelect = itemBlock.querySelector('.item-cara1-panel-type');
        const cara2PanelSelect = itemBlock.querySelector('.item-cara2-panel-type');
        const cieloPanelSelect = itemBlock.querySelector('.item-cielo-panel-type');
        const cenefaPanelSelect = itemBlock.querySelector('.item-cenefa-panel-type'); // New
        if(cara1PanelSelect) populatePanelTypes(cara1PanelSelect);
        if(cara2PanelSelect) populatePanelTypes(cara2PanelSelect);
        if(cieloPanelSelect) populatePanelTypes(cieloPanelSelect);
        if(cenefaPanelSelect) populatePanelTypes(cenefaPanelSelect); // Populate new cenefa select


        // Add event listener to the structure type select element IN THIS BLOCK
        const structureTypeSelect = itemBlock.querySelector('.item-structure-type');
        structureTypeSelect.addEventListener('change', (event) => {
            const selectedType = event.target.value;
            // Actualizamos el título principal del ítem para reflejar el tipo y número.
            updateItemHeaderSummary(itemBlock); // Llama a la función que ahora solo actualiza el h3 con tipo y número.

            // Clear existing segments when changing type
            const muroSegmentsList = itemBlock.querySelector('.muro-segments .segments-list');
            const cieloSegmentsList = itemBlock.querySelector('.cielo-segments .segments-list');
            const cenefaSegmentsList = itemBlock.querySelector('.cenefa-segments .segments-list'); // New


            if (selectedType === 'muro') {
                // Clear cielo and cenefa segments and add a muro segment if needed
                if (cieloSegmentsList) cieloSegmentsList.innerHTML = '';
                if (cenefaSegmentsList) cenefaSegmentsList.innerHTML = ''; // Clear cenefa segments
                if (muroSegmentsList && muroSegmentsList.querySelectorAll('.muro-segment').length === 0) {
                     const newSegment = createMuroSegmentBlock(itemId, 1);
                     muroSegmentsList.appendChild(newSegment);
                     updateSegmentItemSummary(newSegment); // Update summary for newly created segment
                }
            } else if (selectedType === 'cielo') {
                 // Clear muro and cenefa segments and add a cielo segment if needed
                 if (muroSegmentsList) muroSegmentsList.innerHTML = '';
                 if (cenefaSegmentsList) cenefaSegmentsList.innerHTML = ''; // Clear cenefa segments
                 if (cieloSegmentsList && cieloSegmentsList.querySelectorAll('.cielo-segment').length === 0) {
                     const newSegment = createCieloSegmentBlock(itemId, 1);
                     cieloSegmentsList.appendChild(newSegment);
                     updateSegmentItemSummary(newSegment); // Update summary for newly created segment
                 }
            } else if (selectedType === 'cenefa') { // Handle Cenefa type change
                 // Clear muro and cielo segments and add a cenefa segment if needed
                 if (muroSegmentsList) muroSegmentsList.innerHTML = '';
                 if (cieloSegmentsList) cieloSegmentsList.innerHTML = '';
                 if (cenefaSegmentsList && cenefaSegmentsList.querySelectorAll('.cenefa-segment').length === 0) {
                      const newSegment = createCenefaSegmentBlock(itemId, 1); // New function call
                      cenefaSegmentsList.appendChild(newSegment);
                      updateSegmentItemSummary(newSegment); // Update summary for newly created segment
                 }
            }

            updateItemInputVisibility(itemBlock);

            // --- Llama a la función para actualizar el resumen dentro de *todos* los segmentos existentes del ítem después de cambiar el tipo ---
            // Esto asegura que si cambias de muro a cielo, los segmentos existentes (ahora de cielo) actualicen su resumen.
            itemBlock.querySelectorAll('.muro-segment, .cielo-segment, .cenefa-segment').forEach(segBlock => { // Added .cenefa-segment
                updateSegmentItemSummary(segBlock);
            });
            // --- Fin llamado ---

            // Clear results and hide download buttons on type change
             resultsContent.innerHTML = '<p>Tipo de ítem cambiado. Recalcula los materiales totales.</p>';
             downloadOptionsDiv.classList.add('hidden');
             lastCalculatedTotalMaterials = {};
             lastCalculatedItemsSpecs = [];
             lastErrorMessages = [];
             lastCalculatedWorkArea = ''; // Clear stored data on type change
        });

         // --- Agrega event listeners a inputs relevantes del ítem para actualizar el resumen en CADA SEGMENTO ---
         // Estos listeners se agregan al ITEM PADRE, pero iterarán sobre los hijos (segmentos) para actualizar su resumen.
         const relevantInputs = itemBlock.querySelectorAll(
             '.item-structure-type, .item-faces, .item-cara1-panel-type, .item-cara2-panel-type, ' +
             '.item-post-spacing, .item-double-structure, .item-cielo-panel-type, .item-plenum, .item-angular-deduction, ' + // Existing
             '.item-cenefa-orientation, .item-cenefa-panel-type, .item-cenefa-sides, .item-cenefa-liston-spacing' // New Cenefa inputs
         );
         relevantInputs.forEach(input => {
             // Determina el tipo de evento apropiado: 'input' para campos de texto/número, 'change' para selects y checkboxes.
             const eventType = (input.tagName === 'SELECT' || input.type === 'checkbox') ? 'change' : 'input';

             input.addEventListener(eventType, () => {
                 // Cuando un input relevante cambia, actualiza el resumen en TODOS los segmentos de este ítem.
                 itemBlock.querySelectorAll('.muro-segment, .cielo-segment, .cenefa-segment').forEach(segBlock => { // Added .cenefa-segment
                     updateSegmentItemSummary(segBlock);
                 });
                  // Si el input que cambió es el de 'faces', también necesitamos actualizar la visibilidad de los paneles de la Cara 2.
                 if (input.classList.contains('item-faces')) {
                      updateItemInputVisibility(itemBlock); // Esto ya llama a updateSegmentItemSummary dentro si el tipo es muro.
                 }
                 // También actualiza el encabezado principal si es relevante (solo tipo y número)
                 updateItemHeaderSummary(itemBlock);

             });
         });
         // --- Fin Agregación de Event Listeners para Segmentos ---


        // Add event listener to the new remove button
        const removeButton = itemBlock.querySelector('.remove-item-btn');
        removeButton.addEventListener('click', () => {
            itemBlock.remove(); // Remove the block from the DOM
            // Clear results and hide download buttons after removal for immediate feedback
             resultsContent.innerHTML = '<p>Ítem eliminado. Recalcula los materiales totales.</p>';
             downloadOptionsDiv.classList.add('hidden'); // Hide download options
             // Also reset stored data on item removal
             lastCalculatedTotalMaterials = {};
             lastCalculatedItemsSpecs = [];
             lastErrorMessages = [];
             lastCalculatedWorkArea = ''; // Clear stored data on item removal
             // Re-evaluate if calculate button should be disabled (if no items left)
             toggleCalculateButtonState();
        });

        // Set initial visibility for the inputs in the new block (defaults to muro)
        // Esto también llama a updateSegmentItemSummary para los segmentos iniciales si el tipo es muro.
        updateItemInputVisibility(itemBlock);

        // Re-evaluate if calculate button should be enabled (since an item was added)
        toggleCalculateButtonState();

        return itemBlock; // Return the created element
    };

    // --- Function to Enable/Disable Calculate Button ---
    const toggleCalculateButtonState = () => {
        const itemBlocks = itemsContainer.querySelectorAll('.item-block');
        calculateBtn.disabled = itemBlocks.length === 0;
    };


// --- Main Calculation Function for ALL Items ---
const calculateMaterials = () => {
    console.log("Iniciando cálculo de materiales...");
    const itemBlocks = itemsContainer.querySelectorAll('.item-block');

    // --- Lee el valor del nuevo input de Área de Trabajo ---
    const workAreaInput = document.getElementById('work-area');
    const workArea = workAreaInput ? workAreaInput.value.trim() : ''; // Lee el valor y quita espacios al inicio/fin
    console.log(`Área de Trabajo: "${workArea}"`);
    // --- Fin Lectura ---

    // --- Constante añadida para el largo del Canal Soporte ---
    const CANAL_SOPORTE_LARGO_ESTANDAR = 3.66;

    // --- Accumulators for Panels (per panel type) based on Image Logic ---
    let panelAccumulators = {};
    PANEL_TYPES.forEach(type => {
        panelAccumulators[type] = {
            suma_fraccionaria_pequenas: 0.0,
            suma_redondeada_otros: 0
        };
    });
    console.log("Acumuladores de paneles inicializados:", panelAccumulators);

    // --- Accumulator for ALL other materials (rounded per item and summed) ---
    let otherMaterialsTotal = {};
    let currentCalculatedItemsSpecs = []; // Array to store specs of validly calculated items
    let currentErrorMessages = []; // Use an array to collect validation error messages

    // --- Almacena el valor del Área de Trabajo con los resultados actuales ---
    let currentCalculatedWorkArea = workArea;

    // Clear previous results and hide download buttons initially
    resultsContent.innerHTML = '';
    downloadOptionsDiv.classList.add('hidden');

    if (itemBlocks.length === 0) {
        console.log("No hay ítems para calcular.");
        // Uses CSS class for styling
        resultsContent.innerHTML = '<p style="color: orange; text-align: center; font-style: italic;">Por favor, agrega al menos un Muro, Cielo o Cenefa para calcular.</p>'; // Updated message
        // Store empty results
         lastCalculatedTotalMaterials = {}; // No hay totales válidos para descargar
         lastCalculatedItemsSpecs = []; // No hay specs válidos para descargar
         lastErrorMessages = ['No hay ítems agregados para calcular.'];
         lastCalculatedWorkArea = ''; // Limpia el área de trabajo almacenada si no hay cálculo
        return;
    }

    console.log(`Procesando ${itemBlocks.length} ítems.`);
    // Iterate through each item block and calculate its materials
    itemBlocks.forEach(itemBlock => {
         // --- Manejo de Errores a Nivel de Ítem ---
         try {
            const itemNumber = itemBlock.dataset.itemId.split('-')[1]; // Get item number from item ID

            const type = itemBlock.querySelector('.item-structure-type').value;
            const itemId = itemBlock.dataset.itemId; // Get the unique item ID

            // Get common values (some are type-specific and might be NaN/null/false if hidden)
            const facesInput = itemBlock.querySelector('.item-faces');
            const faces = facesInput && !facesInput.closest('.hidden') ? parseInt(facesInput.value) : NaN;

            const plenumInput = itemBlock.querySelector('.item-plenum');
            const plenum = plenumInput && !plenumInput.closest('.hidden') ? parseFloat(plenumInput.value) : NaN;

            // --- Lee el valor del input de descuento angular ---
            const angularDeductionInput = itemBlock.querySelector('.item-angular-deduction');
            const angularDeduction = angularDeductionInput && !angularDeductionInput.closest('.hidden') ? parseFloat(angularDeductionInput.value) : NaN;
            // --- Fin lectura ---

            const isDoubleStructureInput = itemBlock.querySelector('.item-double-structure');
            const isDoubleStructure = isDoubleStructureInput && !isDoubleStructureInput.checked ? false : isDoubleStructureInput.checked;


            const postSpacingInput = itemBlock.querySelector('.item-post-spacing');
            const postSpacing = postSpacingInput && !postSpacingInput.closest('.hidden') ? parseFloat(postSpacingInput.value) : NaN;

            // Get panel types based on visibility and type
            const cara1PanelTypeSelect = itemBlock.querySelector('.item-cara1-panel-type');
            const cara1PanelType = cara1PanelTypeSelect && !cara1PanelTypeSelect.closest('.hidden') ? cara1PanelTypeSelect.value : null;

            const cara2PanelTypeSelect = itemBlock.querySelector('.item-cara2-panel-type');
            // Only read if faces is 2, selector is visible, and the value is not null/empty
            const cara2PanelType = (faces === 2 && cara2PanelTypeSelect && !cara2PanelTypeSelect.closest('.hidden') && cara2PanelTypeSelect.value) ? cara2PanelTypeSelect.value : null;


            const cieloPanelTypeSelect = itemBlock.querySelector('.item-cielo-panel-type');
            const cieloPanelType = cieloPanelTypeSelect && !cieloPanelTypeSelect.closest('.hidden') ? cieloPanelTypeSelect.value : null;

             // Get Cenefa specific values
             const cenefaOrientationSelect = itemBlock.querySelector('.item-cenefa-orientation');
             const cenefaOrientation = cenefaOrientationSelect && !cenefaOrientationSelect.closest('.hidden') ? cenefaOrientationSelect.value : null;

             const cenefaPanelTypeSelect = itemBlock.querySelector('.item-cenefa-panel-type');
             const cenefaPanelType = cenefaPanelTypeSelect && !cenefaPanelTypeSelect.closest('.hidden') ? cenefaPanelTypeSelect.value : null;

             const cenefaSidesInput = itemBlock.querySelector('.item-cenefa-sides');
             const cenefaSides = cenefaSidesInput && !cenefaSidesInput.closest('.hidden') ? parseInt(cenefaSidesInput.value) : NaN;

             const cenefaListonSpacingInput = itemBlock.querySelector('.item-cenefa-liston-spacing');
             const cenefaListonSpacing = cenefaListonSpacingInput && !cenefaListonSpacingInput.closest('.hidden') ? parseFloat(cenefaListonSpacingInput.value) : ESPACIAMIENTO_CANAL_LISTON; // Default if hidden


            console.log(`Procesando Ítem #${itemNumber} (ID: ${itemId}): Tipo=${type}`);

             // Basic Validation for Each Item
             let itemSpecificErrors = [];
             let itemValidatedSpecs = { // Store specs for valid items *before* calculation
                 id: itemId,
                 number: parseInt(itemNumber), // Store number as integer
                 type: type,
                 faces: type === 'muro' ? faces : NaN,   // Only store faces for muros
                 cara1PanelType: type === 'muro' ? cara1PanelType : null, // Only store for muros
                 cara2PanelType: type === 'muro' && faces === 2 ? cara2PanelType : null, // Only store for muros (if 2 faces)
                 cieloPanelType: type === 'cielo' ? cieloPanelType : null, // Only store for cielos
                 cenefaOrientation: type === 'cenefa' ? cenefaOrientation : null, // New: Store for cenefas
                 cenefaPanelType: type === 'cenefa' ? cenefaPanelType : null, // New: Store for cenefas
                 cenefaSides: type === 'cenefa' ? cenefaSides : NaN, // New: Store for cenefas
                 postSpacing: type === 'muro' ? postSpacing : NaN, // Only store for muros
                 plenum: type === 'cielo' ? plenum : NaN, // Used for Cielo
                 angularDeduction: type === 'cielo' ? angularDeduction : NaN, // Store deduction for cielos
                 isDoubleStructure: type === 'muro' ? isDoubleStructure : false, // Only store for muros
                 cenefaListonSpacing: type === 'cenefa' ? cenefaListonSpacing : NaN, // New: Store for cenefas
                 segments: [] // Array to store valid segments (muro, cielo or cenefa)
             };

            // Object to hold calculated *other* materials for THIS single item (initial floats)
            let itemOtherMaterialsFloat = {};


            // --- Calculation Logic for the CURRENT Item ---
            if (type === 'muro') {
                const segmentBlocks = itemBlock.querySelectorAll('.muro-segment');
                let totalMuroAreaForPanelsFinishing = 0; // Renamed for clarity
                let totalMuroWidthForStructure = 0;
                let totalMuroRawWidthForStructure = 0;
                let hasValidSegment = false; // Flag to check if at least one segment is valid

                 if (segmentBlocks.length === 0) {
                     itemSpecificErrors.push('Muro debe tener al menos un segmento de medida.');
                 } else {
                     segmentBlocks.forEach((segBlock, index) => {
                         const segmentWidthRaw = parseFloat(segBlock.querySelector('.item-width').value); // Use .item-width class
                         const segmentHeightRaw = parseFloat(segBlock.querySelector('.item-height').value); // Use .item-height class
                         const segmentNumber = index + 1;

                         if (!isNaN(segmentWidthRaw) && segmentWidthRaw > 0) {
                             totalMuroRawWidthForStructure += segmentWidthRaw;
                         }

                         // Validate segment dimensions
                         if (isNaN(segmentWidthRaw) || segmentWidthRaw <= 0 || isNaN(segmentHeightRaw) || segmentHeightRaw <= 0) {
                             itemSpecificErrors.push(`Segmento ${segmentNumber}: Dimensiones inválidas (Ancho y Alto deben ser > 0)`);
                             return; // Skip this segment but continue validating others
                         }
                         
                         const segmentWidth = applyGeneralRule(segmentWidthRaw);
                         const segmentHeight = applyGeneralRule(segmentHeightRaw);
                         const segmentArea = segmentWidth * segmentHeight;

                         // Criterio de optimización para muros pequeños a 2 caras. Se evalúa con el ancho y alto REAL (raw).
                         if (faces === 2 && segmentWidthRaw <= 0.60 && segmentHeightRaw <= 2.44) {
                             // Si cumple, se asume que 1 solo panel es suficiente para ambas caras.
                             const panelTypeForOptimization = cara1PanelType;
                             if (panelTypeForOptimization) {
                                 // Se suma 1 panel completo al acumulador.
                                 panelAccumulators[panelTypeForOptimization].suma_redondeada_otros += 1;
                                 console.log(`Muro #${itemNumber} Segmento ${segmentNumber} (${panelTypeForOptimization}): Aplicando optimización de 1 panel para 2 caras.`);
                             }
                         } else {
                             // Lógica original para todos los demás casos.
                             // Cálculo para Cara 1
                             const panelTypeFace1 = cara1PanelType;
                             const panelesFloatFace1 = segmentArea / PANEL_RENDIMIENTO_M2;
                             if (segmentArea > 0 && panelTypeFace1) {
                                 if (segmentArea < SMALL_AREA_THRESHOLD_M2) {
                                     panelAccumulators[panelTypeFace1].suma_fraccionaria_pequenas += panelesFloatFace1;
                                 } else {
                                     panelAccumulators[panelTypeFace1].suma_redondeada_otros += roundUpFinalUnit(panelesFloatFace1);
                                 }
                             }

                             // Cálculo para Cara 2
                             if (faces === 2 && cara2PanelType) {
                                 const panelTypeFace2 = cara2PanelType;
                                 const panelesFloatFace2 = segmentArea / PANEL_RENDIMIENTO_M2;
                                 if (segmentArea > 0 && panelTypeFace2) {
                                     if (segmentArea < SMALL_AREA_THRESHOLD_M2) {
                                         panelAccumulators[panelTypeFace2].suma_fraccionaria_pequenas += panelesFloatFace2;
                                     } else {
                                         panelAccumulators[panelTypeFace2].suma_redondeada_otros += roundUpFinalUnit(panelesFloatFace2);
                                     }
                                 }
                             }
                         }

                         // Si el segmento es válido, se procede a sumar sus áreas y anchos para otros cálculos.
                         hasValidSegment = true;
                         totalMuroAreaForPanelsFinishing += segmentArea;
                         totalMuroWidthForStructure += segmentWidth;

                         itemValidatedSpecs.segments.push({
                             number: segmentNumber,
                             width: segmentWidthRaw,
                             height: segmentHeightRaw,
                             area: segmentArea
                         });
                     });

                     if (!hasValidSegment && segmentBlocks.length > 0) itemSpecificErrors.push('Ningún segmento de muro tiene dimensiones válidas (> 0).');
                     if (isNaN(faces) || (faces !== 1 && faces !== 2)) itemSpecificErrors.push('Nº Caras inválido (debe ser 1 o 2)');
                     if (isNaN(postSpacing) || postSpacing <= 0) itemSpecificErrors.push('Espaciamiento Postes inválido (debe ser > 0)');
                     if (!cara1PanelType || !PANEL_TYPES.includes(cara1PanelType)) itemSpecificErrors.push('Tipo de Panel Cara 1 inválido.');
                     if (faces === 2 && itemBlock.querySelector('.cara-2-panel-type-group') && !itemBlock.querySelector('.cara-2-panel-type-group').classList.contains('hidden') && (!cara2PanelType || !PANEL_TYPES.includes(cara2PanelType))) {
                         itemSpecificErrors.push('Tipo de Panel Cara 2 inválido para 2 caras.');
                     }
                     
                     itemValidatedSpecs.totalMuroArea = totalMuroAreaForPanelsFinishing;
                     itemValidatedSpecs.totalMuroWidth = totalMuroWidthForStructure;
                 } 

                if (itemSpecificErrors.length === 0 && hasValidSegment) {
                    
                    let postesFloatSingle = 0;
                    if (postSpacing > 0) {
                        itemValidatedSpecs.segments.forEach(seg => {
                            const rawSegmentWidth = seg.width;
                            const segmentWidthRule = applyGeneralRule(seg.width);
                            const segmentHeightRule = applyGeneralRule(seg.height);
                            if (segmentWidthRule > 0 && postSpacing > 0 && segmentHeightRule > 0) {
                                let postesHorizontalBruto = (rawSegmentWidth > 0 && rawSegmentWidth < postSpacing) ? 2 : Math.floor(segmentWidthRule / postSpacing) + 1;
                                let totalPostesSimpleSegment = (segmentHeightRule <= POSTE_LARGO_ESTANDAR) ? postesHorizontalBruto : (postesHorizontalBruto * (segmentHeightRule + EMPALME_POSTE_LONGITUD)) / POSTE_LARGO_ESTANDAR;
                                postesFloatSingle += totalPostesSimpleSegment;
                            }
                        });
                    }

                    const isFace1Heavy = cara1PanelType === 'Exterior' || cara1PanelType === 'Durock';
                    const isFace2Heavy = cara2PanelType === 'Exterior' || cara2PanelType === 'Durock';

                    if (isDoubleStructure && faces === 2 && isFace1Heavy !== isFace2Heavy) {
                        itemOtherMaterialsFloat['Postes'] = (itemOtherMaterialsFloat['Postes'] || 0) + postesFloatSingle;
                        itemOtherMaterialsFloat['Postes Calibre 20'] = (itemOtherMaterialsFloat['Postes Calibre 20'] || 0) + postesFloatSingle;
                    } else {
                        let finalPostesFloat = postesFloatSingle;
                        if (isDoubleStructure) finalPostesFloat *= 2;
                        const posteType = isFace1Heavy ? 'Postes Calibre 20' : 'Postes';
                        itemOtherMaterialsFloat[posteType] = (itemOtherMaterialsFloat[posteType] || 0) + finalPostesFloat;
                    }

                    const longitudNecesariaSingle = applyGeneralRule(totalMuroWidthForStructure) * 2;
                    const canalesFloatSingle = longitudNecesariaSingle / CANAL_LARGO_ESTANDAR;
                    if (isDoubleStructure && faces === 2 && isFace1Heavy !== isFace2Heavy) {
                        itemOtherMaterialsFloat['Canales'] = (itemOtherMaterialsFloat['Canales'] || 0) + canalesFloatSingle;
                        itemOtherMaterialsFloat['Canales Calibre 20'] = (itemOtherMaterialsFloat['Canales Calibre 20'] || 0) + canalesFloatSingle;
                    } else {
                        let canalesFloat = 0;
                        if (isDoubleStructure && totalMuroRawWidthForStructure < 0.75) {
                            const longitudTotalNecesaria = totalMuroRawWidthForStructure * 4;
                            if (longitudTotalNecesaria <= CANAL_LARGO_ESTANDAR) canalesFloat = 1;
                            else canalesFloat = longitudTotalNecesaria / CANAL_LARGO_ESTANDAR;
                        } else {
                            let finalCanalesFloat = canalesFloatSingle;
                            if (isDoubleStructure) finalCanalesFloat *= 2;
                            canalesFloat = finalCanalesFloat;
                        }
                        const canalType = isFace1Heavy ? 'Canales Calibre 20' : 'Canales';
                        itemOtherMaterialsFloat[canalType] = (itemOtherMaterialsFloat[canalType] || 0) + canalesFloat;
                    }

                    const totalAreaRule = applyGeneralRule(totalMuroAreaForPanelsFinishing);

                    const calculateFinishingForFace = (panelType, area, materialObject) => {
                        if (area <= 0 || !panelType) return;
                        const panelCount = area / PANEL_RENDIMIENTO_M2;
                        if (['Normal', 'Resistente a la Humedad', 'Resistente al Fuego', 'Alta Resistencia'].includes(panelType)) {
                            materialObject['Pasta'] = (materialObject['Pasta'] || 0) + (area / 22);
                            materialObject['Cinta de Papel'] = (materialObject['Cinta de Papel'] || 0) + (area * (7 / PANEL_RENDIMIENTO_M2));
                            materialObject['Lija Grano 120'] = (materialObject['Lija Grano 120'] || 0) + (panelCount / 2);
                            materialObject['Tornillos de 1" punta fina'] = (materialObject['Tornillos de 1" punta fina'] || 0) + (panelCount * 40);
                        } else if (['Exterior', 'Durock'].includes(panelType)) {
                            materialObject['Basecoat'] = (materialObject['Basecoat'] || 0) + (area / 8);
                            materialObject['Cinta malla'] = (materialObject['Cinta malla'] || 0) + (area * 1);
                            materialObject['Tornillos de 1" punta broca'] = (materialObject['Tornillos de 1" punta broca'] || 0) + (panelCount * 40);
                        }
                    };

                    calculateFinishingForFace(cara1PanelType, totalAreaRule, itemOtherMaterialsFloat);
                    if (faces === 2 && cara2PanelType) {
                        calculateFinishingForFace(cara2PanelType, totalAreaRule, itemOtherMaterialsFloat);
                    }

                    if (totalMuroWidthForStructure > 0) {
                        let roundedPostesNormal = roundUpFinalUnit(itemOtherMaterialsFloat['Postes'] || 0);
                        let roundedPostesHeavy = roundUpFinalUnit(itemOtherMaterialsFloat['Postes Calibre 20'] || 0);
                        let roundedCanalesNormal = roundUpFinalUnit(itemOtherMaterialsFloat['Canales'] || 0);
                        let roundedCanalesHeavy = roundUpFinalUnit(itemOtherMaterialsFloat['Canales Calibre 20'] || 0);

                         itemOtherMaterialsFloat['Clavos con Roldana'] = (itemOtherMaterialsFloat['Clavos con Roldana'] || 0) + ((roundedCanalesNormal + roundedCanalesHeavy) * 8);
                         itemOtherMaterialsFloat['Fulminantes'] = (itemOtherMaterialsFloat['Fulminantes'] || 0) + ((roundedCanalesNormal + roundedCanalesHeavy) * 8);
                         itemOtherMaterialsFloat['Tornillos de 1/2" punta fina'] = (itemOtherMaterialsFloat['Tornillos de 1/2" punta fina'] || 0) + (roundedPostesNormal * 4);
                         itemOtherMaterialsFloat['Tornillos de 1/2" punta broca'] = (itemOtherMaterialsFloat['Tornillos de 1/2" punta broca'] || 0) + (roundedPostesHeavy * 4);
                    }
                }

            } else if (type === 'cielo') {
                 const segmentBlocks = itemBlock.querySelectorAll('.cielo-segment');
                 let totalCieloAreaForPanelsFinishing = 0;
                 let totalCieloPerimeterForAngular = 0;
                 let totalLinearMetersSoporte = 0; // Acumulador para Canal Soporte
                 let hasValidSegment = false;
                 let validSegments = [];

                 if (segmentBlocks.length === 0) {
                     itemSpecificErrors.push('Cielo Falso debe tener al menos un segmento de medida.');
                 } else {
                     segmentBlocks.forEach((segBlock, index) => {
                          const segmentWidthRaw = parseFloat(segBlock.querySelector('.item-width').value);
                          const segmentLengthRaw = parseFloat(segBlock.querySelector('.item-length').value);
                          const segmentNumber = index + 1;
                          
                          if (isNaN(segmentWidthRaw) || segmentWidthRaw <= 0 || isNaN(segmentLengthRaw) || segmentLengthRaw <= 0) {
                              itemSpecificErrors.push(`Segmento ${segmentNumber}: Dimensiones inválidas (Ancho y Largo deben ser > 0)`);
                              return;
                          }

                          hasValidSegment = true;
                          
                          // Lógica precisa para Canal Soporte por segmento (usando dimensiones reales)
                          const shorterDim = Math.min(segmentWidthRaw, segmentLengthRaw);
                          const longerDim = Math.max(segmentWidthRaw, segmentLengthRaw);
                          const numChannels = Math.floor(longerDim / ESPACIAMIENTO_CANAL_SOPORTE) + 1;
                          const linearMetersForSegment = numChannels * shorterDim;
                          totalLinearMetersSoporte += linearMetersForSegment;
                          
                          const segmentWidth = applyGeneralRule(segmentWidthRaw);
                          const segmentLength = applyGeneralRule(segmentLengthRaw);
                          const segmentArea = segmentWidth * segmentLength;
                          totalCieloAreaForPanelsFinishing += segmentArea;
                          totalCieloPerimeterForAngular += 2 * (segmentWidth + segmentLength);

                          validSegments.push({ number: segmentNumber, width: segmentWidthRaw, length: segmentLengthRaw, area: segmentArea });
                          
                          const panelTypeCielo = cieloPanelType;
                          const panelesFloatCielo = segmentArea / PANEL_RENDIMIENTO_M2;

                         if (segmentArea > 0 && panelTypeCielo) {
                              if (segmentArea < SMALL_AREA_THRESHOLD_M2) {
                                  panelAccumulators[panelTypeCielo].suma_fraccionaria_pequenas += panelesFloatCielo;
                              } else {
                                  const panelesRoundedCielo = roundUpFinalUnit(panelesFloatCielo);
                                  panelAccumulators[panelTypeCielo].suma_redondeada_otros += panelesRoundedCielo;
                              }
                          }
                     });

                     if (!hasValidSegment && segmentBlocks.length > 0) {
                         itemSpecificErrors.push('Ningún segmento de cielo falso tiene dimensiones válidas (> 0).');
                     }
                     const plenumInput = itemBlock.querySelector('.item-plenum');
                      if (itemBlock.querySelector('.item-plenum-input') && !itemBlock.querySelector('.item-plenum-input').classList.contains('hidden') && (isNaN(plenum) || plenum < 0)) {
                          itemSpecificErrors.push('Pleno inválido (debe ser >= 0)');
                      }
                     if (!cieloPanelType || !PANEL_TYPES.includes(cieloPanelType)) itemSpecificErrors.push('Tipo de Panel de Cielo inválido.');
                      const angularDeductionInput = itemBlock.querySelector('.item-angular-deduction');
                      if (itemBlock.querySelector('.item-angular-deduction-input') && !itemBlock.querySelector('.item-angular-deduction-input').classList.contains('hidden') && (isNaN(angularDeduction) || angularDeduction < 0)) {
                           itemSpecificErrors.push('Metros a descontar de Angular inválido (debe ser >= 0).');
                      }
                      
                     itemValidatedSpecs.segments = validSegments;
                     itemValidatedSpecs.totalCieloArea = totalCieloAreaForPanelsFinishing;
                     itemValidatedSpecs.totalCieloPerimeterSum = totalCieloPerimeterForAngular;
                     itemValidatedSpecs.angularDeduction = angularDeduction;

                 }

                if (itemSpecificErrors.length === 0 && hasValidSegment) {
                     const totalCieloAreaRule = applyGeneralRule(totalCieloAreaForPanelsFinishing);
                     
                     itemOtherMaterialsFloat['Canal Listón'] = (totalCieloAreaRule > 0) ? (totalCieloAreaRule / ESPACIAMIENTO_CANAL_LISTON) / CANAL_LARGO_ESTANDAR : 0;
                     
                     itemOtherMaterialsFloat['Canal Soporte'] = totalLinearMetersSoporte / CANAL_SOPORTE_LARGO_ESTANDAR;

                     const totalCieloPerimeterRule = applyGeneralRule(totalCieloPerimeterForAngular);
                     let adjustedPerimeter = Math.max(0, totalCieloPerimeterRule - (isNaN(angularDeduction) ? 0 : angularDeduction));
                     itemOtherMaterialsFloat['Angular de Lámina'] = adjustedPerimeter / ANGULAR_LARGO_ESTANDAR;

                     let totalPatasBySegments = 0;
                     itemValidatedSpecs.segments.forEach(seg => {
                         const segmentWidthRule = applyGeneralRule(seg.width);
                         const segmentLengthRule = applyGeneralRule(seg.length);
                         if (segmentWidthRule > 0 && segmentLengthRule > 0 && ESPACIAMIENTO_CANAL_SOPORTE > 0) {
                             totalPatasBySegments += Math.floor(segmentLengthRule / ESPACIAMIENTO_CANAL_SOPORTE) * Math.floor(segmentWidthRule / ESPACIAMIENTO_CANAL_SOPORTE);
                         }
                     });
                     itemOtherMaterialsFloat['Patas'] = totalPatasBySegments;

                     const plenumValue = isNaN(plenum) ? 0 : plenum;
                     const roundedPatasForCuelgue = roundUpFinalUnit(itemOtherMaterialsFloat['Patas'] || 0);
                     itemOtherMaterialsFloat['Canal Listón (para cuelgue)'] = (plenumValue > 0 && roundedPatasForCuelgue > 0) ? (roundedPatasForCuelgue * (plenumValue + LONGITUD_EXTRA_PATA)) / CANAL_LARGO_ESTANDAR : 0;

                     const primaryPanelTypeForFinishing = cieloPanelType;
                     const panelesCielo = totalCieloAreaRule / PANEL_RENDIMIENTO_M2;

                     if (['Normal', 'Resistente a la Humedad', 'Resistente al Fuego', 'Alta Resistencia'].includes(primaryPanelTypeForFinishing)) {
                         itemOtherMaterialsFloat['Pasta'] = (itemOtherMaterialsFloat['Pasta'] || 0) + (totalCieloAreaRule > 0 ? totalCieloAreaRule / 22 : 0);
                         itemOtherMaterialsFloat['Cinta de Papel'] = (itemOtherMaterialsFloat['Cinta de Papel'] || 0) + (totalCieloAreaRule > 0 ? totalCieloAreaRule * (7 / PANEL_RENDIMIENTO_M2) : 0);
                         itemOtherMaterialsFloat['Lija Grano 120'] = (itemOtherMaterialsFloat['Lija Grano 120'] || 0) + (panelesCielo > 0 ? panelesCielo / 2 : 0);
                         // CORREGIDO: Añade tornillos de punta fina para paneles estándar
                         itemOtherMaterialsFloat['Tornillos de 1" punta fina'] = (itemOtherMaterialsFloat['Tornillos de 1" punta fina'] || 0) + (panelesCielo > 0 ? panelesCielo * 40 : 0);
                     } else if (['Exterior', 'Durock'].includes(primaryPanelTypeForFinishing)) {
                         itemOtherMaterialsFloat['Basecoat'] = (itemOtherMaterialsFloat['Basecoat'] || 0) + (totalCieloAreaRule > 0 ? totalCieloAreaRule / 8 : 0);
                         itemOtherMaterialsFloat['Cinta malla'] = (itemOtherMaterialsFloat['Cinta malla'] || 0) + (totalCieloAreaRule > 0 ? totalCieloAreaRule * 1 : 0);
                         // CORREGIDO: Añade tornillos de punta broca para paneles tipo Durock/Exterior
                         itemOtherMaterialsFloat['Tornillos de 1" punta broca'] = (itemOtherMaterialsFloat['Tornillos de 1" punta broca'] || 0) + (panelesCielo > 0 ? panelesCielo * 40 : 0);
                     }

                     if (totalCieloAreaRule > 0 || totalCieloPerimeterRule > 0) {
                         let roundedAngularLamina = roundUpFinalUnit(itemOtherMaterialsFloat['Angular de Lámina'] || 0);
                         let roundedCanalSoporte = roundUpFinalUnit(itemOtherMaterialsFloat['Canal Soporte'] || 0);
                         let roundedCanalListon = roundUpFinalUnit(itemOtherMaterialsFloat['Canal Listón'] || 0);
                         let roundedPatas = roundUpFinalUnit(itemOtherMaterialsFloat['Patas'] || 0);
                         let roundedCanalListonCuelgue = roundUpFinalUnit(itemOtherMaterialsFloat['Canal Listón (para cuelgue)'] || 0);

                         itemOtherMaterialsFloat['Clavos con Roldana'] = (itemOtherMaterialsFloat['Clavos con Roldana'] || 0) + (roundedAngularLamina * 5) + (roundedCanalSoporte * 8);
                         itemOtherMaterialsFloat['Fulminantes'] = (itemOtherMaterialsFloat['Fulminantes'] || 0) + (roundedAngularLamina * 5) + (roundedCanalSoporte * 8);
                         itemOtherMaterialsFloat['Tornillos de 1/2" punta fina'] = (itemOtherMaterialsFloat['Tornillos de 1/2" punta fina'] || 0) + (roundedCanalListon * 12) + (roundedPatas * 2) + (roundedCanalListonCuelgue * 2);
                     }
                }

            } else if (type === 'cenefa') {
                 const segmentBlocks = itemBlock.querySelectorAll('.cenefa-segment');
                 let totalCenefaPanelArea = 0;
                 let totalCenefaLargoSum = 0;
                 let totalCenefaAnchoSum = 0;
                 let totalCenefaAltoSum = 0;
                 let hasValidSegment = false;
                 let validSegments = [];

                 const orientation = itemValidatedSpecs.cenefaOrientation;
                 const panelType = itemValidatedSpecs.cenefaPanelType;
                 const sides = itemValidatedSpecs.cenefaSides;
                 const listonSpacing = isNaN(itemValidatedSpecs.cenefaListonSpacing) || itemValidatedSpecs.cenefaListonSpacing <= 0 ? ESPACIAMIENTO_CANAL_LISTON : itemValidatedSpecs.cenefaListonSpacing;


                 if (segmentBlocks.length === 0) {
                      itemSpecificErrors.push('Cenefa debe tener al menos un segmento de medida.');
                 } else {
                      segmentBlocks.forEach((segBlock, index) => {
                           const segmentLargoRaw = parseFloat(segBlock.querySelector('.item-length').value);
                           const segmentAnchoRaw = parseFloat(segBlock.querySelector('.item-width').value);
                           const segmentAltoRaw = parseFloat(segBlock.querySelector('.item-height').value);
                           const segmentNumber = index + 1;

                           const segmentLargo = applyGeneralRule(segmentLargoRaw);
                           const segmentAncho = applyGeneralRule(segmentAnchoRaw);
                           const segmentAlto = applyGeneralRule(segmentAltoRaw);


                           if (isNaN(segmentLargoRaw) || segmentLargoRaw <= 0 || isNaN(segmentAnchoRaw) || segmentAnchoRaw <= 0 || isNaN(segmentAltoRaw) || segmentAltoRaw <= 0) {
                               itemSpecificErrors.push(`Segmento ${segmentNumber}: Dimensiones inválidas (Largo, Ancho y Alto deben ser > 0)`);
                               return;
                           }

                           hasValidSegment = true;

                           let segmentPanelArea = 0;
                           if (orientation === 'horizontal') {
                                segmentPanelArea = segmentLargo * segmentAlto;
                           } else if (orientation === 'vertical') {
                               segmentPanelArea = segmentLargo * segmentAncho;
                           }
                           totalCenefaPanelArea += segmentPanelArea * sides;

                           totalCenefaLargoSum += segmentLargo;
                           totalCenefaAnchoSum += segmentAncho;
                           totalCenefaAltoSum += segmentAlto;


                           validSegments.push({
                              number: segmentNumber,
                              largo: segmentLargoRaw,
                              ancho: segmentAnchoRaw,
                              alto: segmentAltoRaw,
                              panelArea: segmentPanelArea * sides
                          });
                      });

                      if (!hasValidSegment && segmentBlocks.length > 0) {
                         itemSpecificErrors.push('Ningún segmento de cenefa tiene dimensiones válidas (> 0).');
                      }
                      if (!cenefaOrientation || (cenefaOrientation !== 'horizontal' && cenefaOrientation !== 'vertical')) itemSpecificErrors.push('Orientación Principal de Cenefa inválida.');
                      if (!cenefaPanelType || !PANEL_TYPES.includes(cenefaPanelType)) itemSpecificErrors.push('Tipo de Panel de Cenefa inválido.');
                      if (isNaN(cenefaSides) || cenefaSides <= 0) itemSpecificErrors.push('Nº de Lados/Caras de Cenefa inválido (debe ser > 0).');
                       if (isNaN(listonSpacing) || listonSpacing <= 0) {
                           if (itemBlock.querySelector('.item-cenefa-liston-spacing-input') && !itemBlock.querySelector('.item-cenefa-liston-spacing-input').classList.contains('hidden')) {
                                itemSpecificErrors.push('Espaciamiento Canal Listón inválido (debe ser > 0).');
                            }
                       }
                        
                      itemValidatedSpecs.segments = validSegments;
                      itemValidatedSpecs.totalCenefaPanelArea = totalCenefaPanelArea;
                      itemValidatedSpecs.totalCenefaLargoSum = totalCenefaLargoSum;
                      itemValidatedSpecs.totalCenefaAnchoSum = totalCenefaAnchoSum;
                      itemValidatedSpecs.totalCenefaAltoSum = totalCenefaAltoSum;

                 }

                 if (itemSpecificErrors.length === 0 && hasValidSegment) {
                     let panelesCenefaFloat = 0;

                     if (totalCenefaPanelArea > 0) {
                        const panelTypeForScrews = cenefaPanelType;

                        const panelesTeorico = totalCenefaPanelArea / PANEL_RENDIMIENTO_M2;
                        const wasteFactor = 0.15;
                        panelesCenefaFloat = panelesTeorico * (1 + wasteFactor);

                         if (totalCenefaPanelArea > 0 && panelesCenefaFloat <= 1) {
                             panelesCenefaFloat = 1;
                         } else if (panelesCenefaFloat > 1) {
                             panelesCenefaFloat = roundUpFinalUnit(panelesCenefaFloat);
                         } else {
                             panelesCenefaFloat = 0;
                         }

                         if (panelesCenefaFloat > 0 && panelType) {
                              panelAccumulators[panelType].suma_redondeada_otros += panelesCenefaFloat;
                         }

                         const inchScrewType = panelTypeForScrews === 'Exterior' || panelTypeForScrews === 'Durock' ? 'Tornillos de 1" punta broca' : 'Tornillos de 1" punta fina';
                          itemOtherMaterialsFloat[inchScrewType] = (itemOtherMaterialsFloat[inchScrewType] || 0) + (panelesCenefaFloat > 0 ? panelesCenefaFloat * 40 : 0);
                     }

                     let canalListonCenefaFloat = 0;
                     const totalCenefaLargoRule = applyGeneralRule(totalCenefaLargoSum);
                     const totalCenefaAnchoRule = applyGeneralRule(totalCenefaAnchoSum);


                     if (orientation === 'horizontal') {
                          let numberOfLines = totalCenefaLargoRule > 0 && listonSpacing > 0 ? totalCenefaLargoRule / listonSpacing : 0;

                         let effectiveLengthPerLine = 0;
                         if (totalCenefaAnchoRule > 0) {
                             effectiveLengthPerLine = totalCenefaAnchoRule + (Math.ceil(totalCenefaAnchoRule / CANAL_LARGO_ESTANDAR) - 1) * EMPALME_POSTE_LONGITUD;
                             if (effectiveLengthPerLine < 0) effectiveLengthPerLine = 0;
                         }

                         let totalLinearMeters = numberOfLines * effectiveLengthPerLine;

                         if (totalLinearMeters > 0) {
                             canalListonCenefaFloat = totalLinearMeters / CANAL_LARGO_ESTANDAR;
                         }
                         itemOtherMaterialsFloat['Canal Listón (Cenefa Horizontal)'] = canalListonCenefaFloat;
                         
                         const roundedHorizontalBars = roundUpFinalUnit(canalListonCenefaFloat);
                          itemOtherMaterialsFloat['Tornillos de 1/2" punta fina'] = (itemOtherMaterialsFloat['Tornillos de 1/2" punta fina'] || 0) + (roundedHorizontalBars * 4);
                     }

                     let angularLaminaCenefaFloat = 0;

                     if (totalCenefaLargoRule > 0) {
                         const numeroEstimadoInicialDePiezas = Math.ceil(totalCenefaLargoRule / ANGULAR_LARGO_ESTANDAR);
                         const totalEstimatedPieces = numeroEstimadoInicialDePiezas * 2;
                         const cantidadEstimadaDeEmpalmes = Math.max(0, totalEstimatedPieces - 1);
                         const longitudTotalDeMaterialNecesario = (totalCenefaLargoRule * 2) + (cantidadEstimadaDeEmpalmes * EMPALME_ANGULAR_LONGITUD);
                          if (longitudTotalDeMaterialNecesario > 0) {
                             angularLaminaCenefaFloat = longitudTotalDeMaterialNecesario / ANGULAR_LARGO_ESTANDAR;
                         }

                         itemOtherMaterialsFloat['Angular de Lámina (Cenefa)'] = angularLaminaCenefaFloat;
                         
                         const totalAngularLengthCalculated = angularLaminaCenefaFloat * ANGULAR_LARGO_ESTANDAR;
                         const totalAngularLengthRule = applyGeneralRule(totalAngularLengthCalculated);

                         const screwTypeForAngularWall = panelType === 'Exterior' || panelType === 'Durock' ? 'Tornillos de 1" punta broca' : 'Tornillos de 1" punta fina';
                          itemOtherMaterialsFloat[screwTypeForAngularWall] = (itemOtherMaterialsFloat[screwTypeForAngularWall] || 0) + (totalAngularLengthRule > 0 ? totalAngularLengthRule * (5 / ANGULAR_LARGO_ESTANDAR) : 0);
                          
                          const roundedAngularBars = roundUpFinalUnit(angularLaminaCenefaFloat);
                         itemOtherMaterialsFloat['Clavos con Roldana'] = (itemOtherMaterialsFloat['Clavos con Roldana'] || 0) + (roundedAngularBars * 5);
                         itemOtherMaterialsFloat['Fulminantes'] = (itemOtherMaterialsFloat['Fulminantes'] || 0) + (roundedAngularBars * 5);
                     }

                     const primaryPanelTypeForFinishing = cenefaPanelType;
                     const totalCenefaPanelAreaRule = applyGeneralRule(totalCenefaPanelArea);

                      if (['Normal', 'Resistente a la Humedad', 'Resistente al Fuego', 'Alta Resistencia'].includes(primaryPanelTypeForFinishing)) {
                         itemOtherMaterialsFloat['Pasta'] = (itemOtherMaterialsFloat['Pasta'] || 0) + (totalCenefaPanelAreaRule > 0 ? totalCenefaPanelAreaRule / 22 : 0);
                         itemOtherMaterialsFloat['Cinta de Papel'] = (itemOtherMaterialsFloat['Cinta de Papel'] || 0) + (totalCenefaPanelAreaRule > 0 ? totalCenefaPanelAreaRule * (7 / PANEL_RENDIMIENTO_M2) : 0);
                         itemOtherMaterialsFloat['Lija Grano 120'] = (itemOtherMaterialsFloat['Lija Grano 120'] || 0) + (totalCenefaPanelAreaRule > 0 ? (totalCenefaPanelAreaRule / PANEL_RENDIMIENTO_M2) / 2 : 0);
                     } else if (['Exterior', 'Durock'].includes(primaryPanelTypeForFinishing)) {
                         itemOtherMaterialsFloat['Basecoat'] = (itemOtherMaterialsFloat['Basecoat'] || 0) + (totalCenefaPanelAreaRule > 0 ? totalCenefaPanelAreaRule / 8 : 0);
                         itemOtherMaterialsFloat['Cinta malla'] = (itemOtherMaterialsFloat['Cinta malla'] || 0) + (totalCenefaPanelAreaRule > 0 ? totalCenefaPanelAreaRule * 1 : 0);
                     }
                 }
            } else {
                itemSpecificErrors.push('Tipo de estructura desconocido.');
            }

            if (itemSpecificErrors.length > 0) {
                 const errorTitle = `${getItemTypeName(type)} #${itemNumber}`;
                 currentErrorMessages.push(`Error en ${errorTitle}: ${itemSpecificErrors.join(', ')}`);
                 return;
            }
             
             currentCalculatedItemsSpecs.push(itemValidatedSpecs);

            for (const material in itemOtherMaterialsFloat) {
                if (itemOtherMaterialsFloat.hasOwnProperty(material)) {
                    const floatQuantity = itemOtherMaterialsFloat[material];
                    if (!isNaN(floatQuantity)) {
                         const roundedQuantity = roundUpFinalUnit(floatQuantity);
                         otherMaterialsTotal[material] = (otherMaterialsTotal[material] || 0) + roundedQuantity;
                     }
                }
            }
             } catch (error) {
             const itemIdentifier = itemBlock.dataset.itemId ? `#${itemBlock.dataset.itemId.split('-')[1]}` : '(ID desconocido)';
             const itemType = itemBlock.querySelector('.item-structure-type') ? getItemTypeName(itemBlock.querySelector('.item-structure-type').value) : 'Desconocido';
             const errorMessage = `Error inesperado procesando Ítem ${itemType} ${itemIdentifier}: ${error.message}`;
             currentErrorMessages.push(errorMessage);
             console.error(errorMessage, error);
         }
    }); 

    let finalPanelTotals = {};
     for (const type in panelAccumulators) {
         if (panelAccumulators.hasOwnProperty(type)) {
             const acc = panelAccumulators[type];
             const totalPanelsForType = roundUpFinalUnit(acc.suma_fraccionaria_pequenas) + acc.suma_redondeada_otros;
             if (totalPanelsForType > 0) {
                  finalPanelTotals[`Paneles de ${type}`] = totalPanelsForType;
             }
         }
     }

    let finalTotalMaterials = { ...finalPanelTotals, ...otherMaterialsTotal };
    
    if (currentErrorMessages.length > 0) {
         resultsContent.innerHTML = '<div class="error-message"><h2>Errores Encontrados:</h2>' +
                                    currentErrorMessages.map(msg => `<p>${msg}</p>`).join('') +
                                    '<p>Por favor, corrige los errores indicados y vuelve a calcular.</p></div>';
         downloadOptionsDiv.classList.add('hidden');
         lastCalculatedTotalMaterials = {};
         lastCalculatedItemsSpecs = [];
         lastErrorMessages = currentErrorMessages;
         lastCalculatedWorkArea = '';
    }

    if (currentCalculatedItemsSpecs.length > 0) {
        let resultsHtml = resultsContent.innerHTML;
        if (currentErrorMessages.length > 0) resultsHtml += '<hr>';
        
        resultsHtml += '<div class="report-header"><h2>Resumen de Materiales</h2>';
        if (currentCalculatedWorkArea) resultsHtml += `<p><strong>Área de Trabajo:</strong> <span>${currentCalculatedWorkArea}</span></p>`;
        resultsHtml += `<p>Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}</p></div><hr>`;
        resultsHtml += '<h3>Detalle de Ítems Calculados:</h3>';

        currentCalculatedItemsSpecs.forEach(item => {
            resultsHtml += `<div class="item-summary"><h4>${getItemTypeName(item.type)} #${item.number}</h4>`;
            resultsHtml += `<p><strong>Tipo:</strong> <span>${getItemTypeName(item.type)}</span></p>`;
            if (item.type === 'muro') {
                if (!isNaN(item.faces)) resultsHtml += `<p><strong>Nº Caras:</strong> <span>${item.faces}</span></p>`;
                if (item.cara1PanelType) resultsHtml += `<p><strong>Panel Cara 1:</strong> <span>${item.cara1PanelType}</span></p>`;
                if (item.faces === 2 && item.cara2PanelType) resultsHtml += `<p><strong>Panel Cara 2:</strong> <span>${item.cara2PanelType}</span></p>`;
                if (!isNaN(item.postSpacing)) resultsHtml += `<p><strong>Espaciamiento Postes:</strong> <span>${item.postSpacing.toFixed(2)} m</span></p>`;
                resultsHtml += `<p><strong>Estructura Doble:</strong> <span>${item.isDoubleStructure ? 'Sí' : 'No'}</span></p>`;
            } else if (item.type === 'cielo') {
                 if (item.cieloPanelType) resultsHtml += `<p><strong>Tipo de Panel:</strong> <span>${item.cieloPanelType}</span></p>`;
                 if (!isNaN(item.plenum)) resultsHtml += `<p><strong>Pleno:</strong> <span>${item.plenum.toFixed(2)} m</span></p>`;
            } else if (item.type === 'cenefa') {
                 if (item.cenefaOrientation) resultsHtml += `<p><strong>Orientación:</strong> <span>${item.cenefaOrientation}</span></p>`;
            }
            resultsHtml += `</div>`;
        });
        resultsHtml += '<hr>';

        resultsHtml += '<h3>Totales de Materiales (Cantidades a Comprar):</h3>';
        const sortedMaterials = Object.keys(finalTotalMaterials).sort();
        resultsHtml += '<table><thead><tr><th>Material</th><th>Cantidad</th><th>Unidad</th></tr></thead><tbody>';

       if (sortedMaterials.length > 0) {
            sortedMaterials.forEach(material => {
                 resultsHtml += `<tr><td>${material}</td><td>${finalTotalMaterials[material]}</td><td>${getMaterialUnit(material)}</td></tr>`;
            });
            downloadOptionsDiv.classList.remove('hidden');
       } else {
            resultsHtml += '<tr><td colspan="3">No se calcularon materiales totales.</td></tr>';
            downloadOptionsDiv.classList.add('hidden');
       }
       resultsHtml += '</tbody></table>';
        resultsContent.innerHTML = resultsHtml;

        lastCalculatedTotalMaterials = finalTotalMaterials;
        lastCalculatedItemsSpecs = currentCalculatedItemsSpecs;
        lastCalculatedWorkArea = currentCalculatedWorkArea;
     } else if (currentErrorMessages.length === 0) {
          resultsContent.innerHTML += '<p>No se pudieron calcular los materiales. Revisa las dimensiones de tus segmentos.</p>';
          downloadOptionsDiv.classList.add('hidden');
          lastCalculatedTotalMaterials = {};
          lastCalculatedItemsSpecs = [];
          lastErrorMessages = [];
          lastCalculatedWorkArea = '';
    }
};

    // --- PDF Generation Function ---
    // Requires the jspdf and jspdf-autotable libraries to be included in index.html
    const generatePDF = () => {
        console.log("Iniciando generación de PDF...");
        // Ensure there are calculated results to download
       if (Object.keys(lastCalculatedTotalMaterials).length === 0 || lastCalculatedItemsSpecs.length === 0) {
           console.warn("No hay resultados calculados para generar el PDF.");
           alert("Por favor, realiza un cálculo válido antes de generar el PDF.");
           return;
       }

       // Initialize jsPDF
       const { jsPDF } = window.jspdf; // Assumes jspdf is loaded globally
       const doc = new jsPDF();

       // Define colors in RGB from CSS variables (using approximations based on common web colors)
       // These should ideally match the CSS variables for consistency in the PDF report styling
       const primaryOliveRGB = [85, 107, 47]; // #556B2F
       const secondaryOliveRGB = [128, 128, 0]; // #808000 (Approximation, could be adjusted to match CSS more closely)
       const darkGrayRGB = [51, 51, 51]; // #333
       const mediumGrayRGB = [102, 102, 102]; // #666
       const lightGrayRGB = [224, 224, 224]; // #e0e0e0
       const extraLightGrayRGB = [248, 248, 248]; // #f8f8f8


       // --- Add Header ---
       doc.setFontSize(18);
       doc.setTextColor(primaryOliveRGB[0], primaryOliveRGB[1], primaryOliveRGB[2]);
       doc.setFont("helvetica", "bold"); // Use a standard font or include custom fonts
       doc.text("Resumen de Materiales Tablayeso", 14, 22);

       doc.setFontSize(10);
       doc.setTextColor(mediumGrayRGB[0], mediumGrayRGB[1], mediumGrayRGB[2]);
       doc.setFont("helvetica", "normal");
       doc.text(`Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}`, 14, 28);

        // --- Agrega el Área de Trabajo al Encabezado del PDF ---
        let currentY = 35; // Posición inicial después de la fecha
        if (lastCalculatedWorkArea) {
             doc.text(`Área de Trabajo: ${lastCalculatedWorkArea}`, 14, currentY);
             currentY += 7; // Deja un poco de espacio si se muestra el área de trabajo
        }
        // Set starting Y position for the next content block
        let finalY = currentY;
        // --- Fin Agregar Área de Trabajo ---


       // --- Add Item Summaries ---
       if (lastCalculatedItemsSpecs.length > 0) {
            console.log("Añadiendo resumen de ítems al PDF.");
            doc.setFontSize(14);
            doc.setTextColor(secondaryOliveRGB[0], secondaryOliveRGB[1], secondaryOliveRGB[2]);
           doc.setFont("helvetica", "bold");
            doc.text("Detalle de Ítems Calculados:", 14, finalY + 10);
            finalY += 15; // Move Y below the title

           const itemSummaryLineHeight = 5; // Space between summary lines within an item
           const itemBlockSpacing = 8; // Space between different item summaries

            lastCalculatedItemsSpecs.forEach(item => {
                // Add item title (using type and number from specs)
                doc.setFontSize(10);
                doc.setTextColor(primaryOliveRGB[0], primaryOliveRGB[1], primaryOliveRGB[2]);
                doc.setFont("helvetica", "bold");
                doc.text(`${getItemTypeName(item.type)} #${item.number}:`, 14, finalY + itemSummaryLineHeight);
                finalY += itemSummaryLineHeight * 1.5; // Move down after the title

                // Add general item details (indented)
                doc.setFontSize(9);
                doc.setTextColor(darkGrayRGB[0], darkGrayRGB[1], darkGrayRGB[2]);
                doc.setFont("helvetica", "normal");

                doc.text(`Tipo: ${getItemTypeName(item.type)}`, 20, finalY + itemSummaryLineHeight);
                finalY += itemSummaryLineHeight;

                // Add type-specific details
                if (item.type === 'muro') {
                     if (!isNaN(item.faces)) {
                          doc.text(`Nº Caras: ${item.faces}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                     if (item.cara1PanelType) {
                          doc.text(`Panel Cara 1: ${item.cara1PanelType}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                     if (item.faces === 2 && item.cara2PanelType) {
                          doc.text(`Panel Cara 2: ${item.cara2PanelType}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                      if (!isNaN(item.postSpacing)) {
                         doc.text(`Espaciamiento Postes: ${item.postSpacing.toFixed(2)} m`, 20, finalY + itemSummaryLineHeight);
                         finalY += itemSummaryLineHeight;
                     }
                      doc.text(`Estructura Doble: ${item.isDoubleStructure ? 'Sí' : 'No'}`, 20, finalY + itemSummaryLineHeight);
                      finalY += itemSummaryLineHeight;

                     doc.text(`Segmentos:`, 20, finalY + itemSummaryLineHeight);
                     finalY += itemSummaryLineHeight;
                     if (item.segments && item.segments.length > 0) {
                         item.segments.forEach(seg => {
                            doc.text(`- Segmento ${seg.number}: ${seg.width.toFixed(2)} m (Ancho) x ${seg.height.toFixed(2)} m (Alto)`, 25, finalY + itemSummaryLineHeight);
                            finalY += itemSummaryLineHeight;
                         });
                         if (!isNaN(item.totalMuroArea)) {
                              doc.text(`- Área Total Segmentos: ${item.totalMuroArea.toFixed(2)} m²`, 25, finalY + itemSummaryLineHeight);
                              finalY += itemSummaryLineHeight;
                          }
                           if (!isNaN(item.totalMuroWidth)) {
                              doc.text(`- Ancho Total Segmentos: ${item.totalMuroWidth.toFixed(2)} m`, 25, finalY + itemSummaryLineHeight);
                              finalY += itemSummaryLineHeight;
                          }

                      } else {
                           doc.text(`- Sin segmentos válidos`, 25, finalY + itemSummaryLineHeight);
                           finalY += itemSummaryLineHeight;
                      }


                } else if (item.type === 'cielo') {
                    if (item.cieloPanelType) {
                          doc.text(`Tipo de Panel: ${item.cieloPanelType}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                    if (!isNaN(item.plenum)) {
                        doc.text(`Pleno: ${item.plenum.toFixed(2)} m`, 20, finalY + itemSummaryLineHeight);
                        finalY += itemSummaryLineHeight;
                    }
                    // --- Agrega el Descuento Angular al Resumen del PDF ---
                    // Asegúrate de que el valor sea un número válido antes de mostrar
                    if (!isNaN(item.angularDeduction) && item.angularDeduction > 0) { // Muestra solo si es > 0 para claridad
                        doc.text(`Descuento Angular: ${item.angularDeduction.toFixed(2)} m`, 20, finalY + itemSummaryLineHeight);
                        finalY += itemSummaryLineHeight;
                    }
                    // --- Fin Agregación ---

                     doc.text(`Segmentos:`, 20, finalY + itemSummaryLineHeight);
                     finalY += itemSummaryLineHeight;
                     if (item.segments && item.segments.length > 0) {
                         item.segments.forEach(seg => {
                             doc.text(`- Segmento ${seg.width.toFixed(2)} m (Ancho) x ${seg.length.toFixed(2)} m (Largo)`, 25, finalY + itemSummaryLineHeight);
                             finalY += itemSummaryLineHeight;
                         });
                         if (!isNaN(item.totalCieloArea)) {
                             doc.text(`- Área Total Segmentos: ${item.totalCieloArea.toFixed(2)} m²`, 25, finalY + itemSummaryLineHeight);
                             finalY += itemSummaryLineHeight;
                         }
                          if (!isNaN(item.totalCieloPerimeterSum)) {
                             // Nota: totalCieloPerimeterSum ahora es la suma de perímetros completos
                             doc.text(`- Suma Perímetros Segmentos: ${item.totalCieloPerimeterSum.toFixed(2)} m`, 25, finalY + itemSummaryLineHeight);
                             finalY += itemSummaryLineHeight;
                         }
                     } else {
                         doc.text(`- Sin segmentos válidos`, 25, finalY + itemSummaryLineHeight);
                         finalY += itemSummaryLineHeight;
                     }
                } else if (item.type === 'cenefa') { // Display Cenefa summary details in PDF
                     if (item.cenefaOrientation) {
                          doc.text(`Orientación Principal: ${item.cenefaOrientation}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                     if (item.cenefaPanelType) {
                          doc.text(`Tipo de Panel: ${item.cenefaPanelType}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                     if (!isNaN(item.cenefaSides)) {
                          doc.text(`Nº de Lados/Caras con Panel: ${item.cenefaSides}`, 20, finalY + itemSummaryLineHeight);
                          finalY += itemSummaryLineHeight;
                     }
                      if (!isNaN(item.cenefaListonSpacing)) {
                         doc.text(`Espaciamiento Canal Listón: ${item.cenefaListonSpacing.toFixed(2)} m`, 20, finalY + itemSummaryLineHeight);
                         finalY += itemSummaryLineHeight;
                      }


                     doc.text(`Segmentos:`, 20, finalY + itemSummaryLineHeight);
                     finalY += itemSummaryLineHeight;
                     if (item.segments && item.segments.length > 0) {
                         item.segments.forEach(seg => {
                             doc.text(`- Segmento ${seg.number}: ${seg.largo.toFixed(2)} m (Largo) x ${seg.ancho.toFixed(2)} m (Ancho) x ${seg.alto.toFixed(2)} m (Alto)`, 25, finalY + itemSummaryLineHeight);
                             finalY += itemSummaryLineHeight;
                         });
                          if (!isNaN(item.totalCenefaPanelArea)) {
                             doc.text(`- Área Total Panel (con lados): ${item.totalCenefaPanelArea.toFixed(2)} m²`, 25, finalY + itemSummaryLineHeight);
                             finalY += itemSummaryLineHeight;
                          }
                          if (!isNaN(item.totalCenefaLargoSum)) {
                             doc.text(`- Suma Largo Segmentos: ${item.totalCenefaLargoSum.toFixed(2)} m`, 25, finalY + itemSummaryLineHeight);
                             finalY += itemSummaryLineHeight;
                          }
                          if (!isNaN(item.totalCenefaAnchoSum)) {
                             doc.text(`- Suma Ancho Segmentos: ${item.totalCenefaAnchoSum.toFixed(2)} m`, 25, finalY + itemSummaryLineHeight);
                             finalY += itemSummaryLineHeight;
                          }
                           if (!isNaN(item.totalCenefaAltoSum)) {
                             doc.text(`- Suma Alto Segmentos: ${item.totalCenefaAltoSum.toFixed(2)} m`, 25, finalY + itemSummaryLineHeight);
                             finalY += itemSummaryLineHeight;
                          }

                     } else {
                         doc.text(`- Sin segmentos válidos`, 25, finalY + itemSummaryLineHeight);
                         finalY += itemSummaryLineHeight;
                     }

                }
                finalY += itemBlockSpacing; // Add space after each item summary block
            });
            finalY += 5; // Add space before the total materials table title
       } else {
            console.log("No hay ítems calculados válidamente para añadir resumen al PDF.");
       }


       // --- Add Total Materials Table ---
        console.log("Añadiendo tabla de materiales totales al PDF.");
       doc.setFontSize(14);
       doc.setTextColor(secondaryOliveRGB[0], secondaryOliveRGB[1], secondaryOliveRGB[2]);
       doc.setFont("helvetica", "bold");
       doc.text("Totales de Materiales:", 14, finalY + 10);
       finalY += 15; // Move Y below the title

       const tableColumn = ["Material", "Cantidad", "Unidad"];
       const tableRows = [];

       // Prepare data for the table
       const sortedMaterials = Object.keys(lastCalculatedTotalMaterials).sort();
       sortedMaterials.forEach(material => {
           const cantidad = lastCalculatedTotalMaterials[material];
           const unidad = getMaterialUnit(material); // Get unit using the helper function
           // Use the material name directly from the key
            tableRows.push([material, cantidad, unidad]);
       });

        // Add the table using jspdf-autotable (requires jspdf-autotable library)
        // Uses color definitions to match the CSS theme
        doc.autoTable({
            head: [tableColumn],
            body: tableRows,
            startY: finalY, // Start position below the last content
            theme: 'plain', // Start with a plain theme to apply custom styles
            headStyles: {
                fillColor: lightGrayRGB, // Use light gray for header background
                textColor: darkGrayRGB, // Use dark gray for header text
                fontStyle: 'bold',
                halign: 'center', // Horizontal alignment
                valign: 'middle', // Vertical alignment
                lineWidth: 0.1, // Add border to cells
                lineColor: lightGrayRGB, // Border color same as background for subtle effect
                fontSize: 10 // Match HTML table header font size (approx)
            },
            bodyStyles: {
                textColor: darkGrayRGB, // Use dark gray for body text
                lineWidth: 0.1, // Add border to cells
                lineColor: lightGrayRGB, // Border color same as header background
                fontSize: 9 // Match HTML table body font size (approx)
            },
             alternateRowStyles: { // Styling for alternate rows
                fillColor: extraLightGrayRGB, // Use extra light gray for alternate rows
            },
             // Specific column styles (Cantidad column is the second one, index 1)
            columnStyles: {
                1: {
                    halign: 'right', // Align quantity to the right
                    fontStyle: 'bold',
                    textColor: primaryOliveRGB // Use primary olive for quantity text color
                },
                 2: { // Unit column
                    halign: 'center' // Align unit to the center or left as preferred
                }
            },
            margin: { top: 10, right: 14, bottom: 14, left: 14 }, // Add margin
             didDrawPage: function (data) {
               // Optional: Add page number or footer here
               doc.setFontSize(8);
               doc.setTextColor(mediumGrayRGB[0], mediumGrayRGB[1], mediumGrayRGB[2]);
               // Using `doc.internal.pageSize.getWidth()` to center the footer roughly
               const footerText = '© 2025 PROPUL - Calculadora de Materiales Tablayeso v2.0'; // Replaced placeholder
               const textWidth = doc.getStringUnitWidth(footerText) * doc.internal.getFontSize() / doc.internal.scaleFactor;
               const centerX = (doc.internal.pageSize.getWidth() - textWidth) / 2;
               doc.text(footerText, centerX, doc.internal.pageSize.height - 10);

               // Add simple page number
               const pageNumberText = `Página ${data.pageNumber}`;
               const pageNumberWidth = doc.getStringUnitWidth(pageNumberText) * doc.internal.getFontSize() / doc.internal.scaleFactor;
               const pageNumberX = doc.internal.pageSize.getWidth() - data.settings.margin.right - pageNumberWidth;
               doc.text(pageNumberText, pageNumberX, doc.internal.pageSize.height - 10);
            }
        });
       // Update finalY after the table
       finalY = doc.autoTable.previous.finalY;

       console.log("PDF generado.");

       // --- Save the PDF ---
       doc.save(`Calculo_Materiales_${new Date().toLocaleDateString('es-ES').replace(/\//g, '-')}.pdf`); // Filename with date
   };


// --- Excel Generation Function ---
// Requires the xlsx library to be included in index.html
const generateExcel = () => {
    console.log("Iniciando generación de Excel...");
    // Ensure there are calculated results to download
   if (Object.keys(lastCalculatedTotalMaterials).length === 0 || lastCalculatedItemsSpecs.length === 0) {
       console.warn("No hay resultados calculados para generar el Excel.");
       alert("Por favor, realiza un cálculo válido antes de generar el Excel.");
       return;
   }

    // Assumes you have loaded the xlsx library globally via a script tag
   if (typeof XLSX === 'undefined') {
        console.error("La librería xlsx no está cargada.");
        alert("Error al generar Excel: Librería xlsx no encontrada.");
        return;
   }


   // Data for the Excel sheet (array of arrays)
   let sheetData = [];

   // Add Header
   sheetData.push(["Calculadora de Materiales Tablayeso"]);
   sheetData.push([`Fecha del cálculo: ${new Date().toLocaleDateString('es-ES')}`]);
    // --- Agrega el Área de Trabajo al Encabezado del Excel ---
   if (lastCalculatedWorkArea) {
       sheetData.push([`Área de Trabajo: ${lastCalculatedWorkArea}`]);
   }
   // --- Fin Agregar Área de Trabajo ---
   sheetData.push([]); // Fila en blanco para espaciar

   // Add Item Summaries
    console.log("Añadiendo resumen de ítems al Excel.");
    sheetData.push(["Detalle de Ítems Calculados:"]);
   // --- ENCABEZADOS DE LA TABLA DE DETALLE DE ÍTEMS ---
   // Se incluye la columna 'Metros Descuento Angular (m)' después de 'Pleno (m)'
   sheetData.push([
       "Tipo Item", "Número Item", "Detalle/Dimensiones",
       "Nº Caras", "Panel Cara 1", "Panel Cara 2", "Tipo Panel Cielo", "Tipo Panel Cenefa", // Added Type Panel Cenefa
       "Espaciamiento Postes (m)", "Pleno (m)", "Metros Descuento Angular (m)", "Estructura Doble", // Existing
       "Orientación Cenefa", "Lados/Caras Cenefa", "Esp. Listón Cenefa (m)", "Esp. Soporte Cenefa (m)", // New Cenefa Config
       "Suma Perímetros Segmentos (m)", "Ancho Total (muro) (m)", "Área Total (m²)", "Suma Largo Segmentos (m)", "Suma Ancho Segmentos (m)", "Suma Alto Segmentos (m)", "Área Total Panel Cenefa (m²)" // Total Columns
   ]);
   // --- FIN ENCABEZADOS ---


   lastCalculatedItemsSpecs.forEach(item => {
        if (item.type === 'muro') {
            // --- Detalles comunes a este ítem Muro (configuración) ---
            // Aseguramos que el array tenga el tamaño correcto, llenando las columnas de Cielo y Cenefa con vacío.
            const muroCommonDetails = [
                getItemTypeName(item.type), // 0: Tipo Item
                item.number,                 // 1: Número Item
                '',                          // 2: Placeholder para Detalle/Dimensiones
                !isNaN(item.faces) ? item.faces : '', // 3: Nº Caras
                item.cara1PanelType ? item.cara1PanelType : '', // 4: Panel Cara 1
                item.faces === 2 && item.cara2PanelType ? item.cara2PanelType : '', // 5: Panel Cara 2
                '',                                          // 6: Tipo Panel Cielo (vacío para muro)
                '',                                          // 7: Tipo Panel Cenefa (vacío para muro)
                !isNaN(item.postSpacing) ? item.postSpacing.toFixed(2) : '', // 8: Espaciamiento Postes
                '',                                          // 9: Pleno (vacío para muro)
                '',                                          // 10: Metros Descuento Angular (vacío para muro)
                item.isDoubleStructure ? 'Sí' : 'No', // 11: Estructura Doble
                // New Cenefa Config Columns (empty for muro)
                '', '', '', '',
                // Total Columns
               '',                                                // Suma Perímetros Segmentos (vacío para Muro)
               !isNaN(item.totalMuroWidth) ? item.totalMuroWidth.toFixed(2) : '', // Ancho Total (m) del Muro
               !isNaN(item.totalMuroArea) ? item.totalMuroArea.toFixed(2) : '',     // Área Total (m²) del Muro
               '', '', '', '' // Suma Largo, Ancho, Alto Segmentos Cenefa, Área Total Panel Cenefa (vacío para Muro)
            ];
            // Fila principal con opciones y totales del ítem
            const muroSummaryRow = [...muroCommonDetails]; // Copia los detalles comunes
            muroSummaryRow[2] = 'Opciones:'; // Etiqueta en la columna de detalle
            sheetData.push(muroSummaryRow);

            // Fila que etiqueta la sección de Segmentos
             const muroSegmentsLabelRow = [...muroCommonDetails];
             muroSegmentsLabelRow[2] = 'Segmentos:'; // Etiqueta "Segmentos:"
              // Agrega celdas vacías para las columnas de totales
              muroSegmentsLabelRow[muroCommonDetails.length] = ''; // Suma Perímetros Segmentos
              muroSegmentsLabelRow[muroCommonDetails.length + 1] = ''; // Ancho Total (muro)
              muroSegmentsLabelRow[muroCommonDetails.length + 2] = ''; // Área Total (m²)
              // Add empty cells for new Cenefa total columns
              muroSegmentsLabelRow[muroCommonDetails.length + 3] = ''; // Suma Largo
              muroSegmentsLabelRow[muroCommonDetails.length + 4] = ''; // Suma Ancho
              muroSegmentsLabelRow[muroCommonDetails.length + 5] = ''; // Suma Alto
              muroSegmentsLabelRow[muroCommonDetails.length + 6] = ''; // Área Total Panel Cenefa

             sheetData.push(muroSegmentsLabelRow);

            if (item.segments && item.segments.length > 0) {
                 item.segments.forEach(seg => {
                      // Fila para cada Segmento individual
                      const segmentRow = [...muroCommonDetails]; // Copia los detalles comunes para esta fila
                      segmentRow[2] = `- Seg ${seg.number}: ${seg.width.toFixed(2)}m x ${seg.height.toFixed(2)}m`; // Dimensiones del segmento
                      // Agrega celdas vacías para las columnas de totales (repetir totales vacíos para cada fila de segmento)
                       segmentRow[muroCommonDetails.length] = ''; // Suma Perímetros Segmentos
                       segmentRow[muroCommonDetails.length + 1] = ''; // Ancho Total (muro)
                       segmentRow[muroCommonDetails.length + 2] = ''; // Área Total (m²)
                       // Add empty cells for new Cenefa total columns
                       segmentRow[muroCommonDetails.length + 3] = ''; // Suma Largo
                       segmentRow[muroCommonDetails.length + 4] = ''; // Suma Ancho
                       segmentRow[muroCommonDetails.length + 5] = ''; // Suma Alto
                       segmentRow[muroCommonDetails.length + 6] = ''; // Área Total Panel Cenefa

                       sheetData.push(segmentRow);
                 });
            } else {
                   // Fila para "Sin segmentos válidos"
                   const noSegmentsRow = [...muroCommonDetails];
                   noSegmentsRow[2] = `- Sin segmentos válidos`; // Mensaje
                    // Agrega celdas vacías para las columnas de totales
                    noSegmentsRow[muroCommonDetails.length] = ''; // Suma Perímetros Segmentos
                    noSegmentsRow[muroCommonDetails.length + 1] = ''; // Ancho Total (muro)
                    noSegmentsRow[muroCommonDetails.length + 2] = ''; // Área Total (m²)
                    // Add empty cells for new Cenefa total columns
                    noSegmentsRow[muroCommonDetails.length + 3] = ''; // Suma Largo
                    noSegmentsRow[muroCommonDetails.length + 4] = ''; // Suma Ancho
                    noSegmentsRow[muroCommonDetails.length + 5] = ''; // Suma Alto
                    noSegmentsRow[muroCommonDetails.length + 6] = ''; // Área Total Panel Cenefa

                    sheetData.push(noSegmentsRow);
            }


        } else if (item.type === 'cielo') {
            // --- Detalles comunes a este ítem Cielo (configuración) ---
            // Aseguramos que el array tenga el tamaño correcto, llenando las columnas de Muro y Cenefa con vacío.
            const cieloCommonDetails = [
                getItemTypeName(item.type), // 0: Tipo Item
                item.number,                 // 1: Número Item
                '',                          // 2: Placeholder para Detalle/Dimensiones
                '',                          // 3: Nº Caras (vacío para cielo)
                '', '',                     // 4, 5: Panel Cara 1, Cara 2 (vacío para cielo)
                item.cieloPanelType ? item.cieloPanelType : '', // 6: Tipo Panel Cielo
                '',                                          // 7: Tipo Panel Cenefa (vacío para cielo)
                '',                          // 8: Espaciamiento Postes (vacío para cielo)
                !isNaN(item.plenum) ? item.plenum.toFixed(2) : '', // 9: Pleno
                !isNaN(item.angularDeduction) ? item.angularDeduction.toFixed(2) : '', // 10: Metros Descuento Angular
                '',                           // 11: Estructura Doble (vacío para cielo)
                // New Cenefa Config Columns (empty for cielo)
                '', '', '', '',
                 // Total Columns
               !isNaN(item.totalCieloPerimeterSum) ? item.totalCieloPerimeterSum.toFixed(2) : '', // Suma Perímetros Segmentos
               '',                                                                                 // Ancho Total (muro) - Vacío para Cielo
               !isNaN(item.totalCieloArea) ? item.totalCieloArea.toFixed(2) : '', // Área Total (m²) del Cielo
                '', '', '', '' // Suma Largo, Ancho, Alto Segmentos Cenefa, Área Total Panel Cenefa (vacío para Cielo)
            ];
            // Fila principal con opciones y totales del ítem
            const cieloSummaryRow = [...cieloCommonDetails]; // Copia los detalles comunes
            cieloSummaryRow[2] = 'Opciones:'; // Etiqueta en la columna de detalle
            sheetData.push(cieloSummaryRow);

            // Fila que etiqueta la sección de Segmentos
             const cieloSegmentsLabelRow = [...cieloCommonDetails];
             cieloSegmentsLabelRow[2] = 'Segmentos:'; // Etiqueta "Segmentos:"
              // Agrega celdas vacías para las columnas de totales
             cieloSegmentsLabelRow[cieloCommonDetails.length] = ''; // Suma Perímetros Segmentos
             cieloSegmentsLabelRow[cieloCommonDetails.length + 1] = ''; // Ancho Total (muro)
             cieloSegmentsLabelRow[cieloCommonDetails.length + 2] = ''; // Área Total (m²)
              // Add empty cells for new Cenefa total columns
             cieloSegmentsLabelRow[cieloCommonDetails.length + 3] = ''; // Suma Largo
             cieloSegmentsLabelRow[cieloCommonDetails.length + 4] = ''; // Suma Ancho
             cieloSegmentsLabelRow[cieloCommonDetails.length + 5] = ''; // Suma Alto
             cieloSegmentsLabelRow[cieloCommonDetails.length + 6] = ''; // Área Total Panel Cenefa

              sheetData.push(cieloSegmentsLabelRow);

            if (item.segments && item.segments.length > 0) {
                  item.segments.forEach(seg => {
                      // Fila para cada Segmento individual
                      const segmentRow = [...cieloCommonDetails]; // Copia los detalles comunes para esta fila
                      segmentRow[2] = `Seg ${seg.number}: ${seg.width.toFixed(2)}m x ${seg.length.toFixed(2)}m`; // Dimensiones del segmento
                       // Agrega celdas vacías para las columnas de totales (repetir totales vacíos para cada fila de segmento)
                       segmentRow[cieloCommonDetails.length] = ''; // Suma Perímetros Segmentos
                       segmentRow[cieloCommonDetails.length + 1] = ''; // Ancho Total (muro)
                       segmentRow[cieloCommonDetails.length + 2] = ''; // Área Total (m²)
                        // Add empty cells for new Cenefa total columns
                       segmentRow[cieloCommonDetails.length + 3] = ''; // Suma Largo
                       segmentRow[cieloCommonDetails.length + 4] = ''; // Suma Ancho
                       segmentRow[cieloCommonDetails.length + 5] = ''; // Suma Alto
                       segmentRow[cieloCommonDetails.length + 6] = ''; // Área Total Panel Cenefa

                       sheetData.push(segmentRow);
                  });
            } else {
                   // Fila para "Sin segmentos válidos"
                   const noSegmentsRow = [...cieloCommonDetails];
                   noSegmentsRow[2] = `- Sin segmentos válidos`; // Mensaje
                    // Agrega celdas vacías para las columnas de totales
                   noSegmentsRow[cieloCommonDetails.length] = ''; // Suma Perímetros Segmentos
                   noSegmentsRow[cieloCommonDetails.length + 1] = ''; // Ancho Total (muro)
                   noSegmentsRow[cieloCommonDetails.length + 2] = ''; // Área Total (m²)
                    // Add empty cells for new Cenefa total columns
                   noSegmentsRow[cieloCommonDetails.length + 3] = ''; // Suma Largo
                   noSegmentsRow[cieloCommonDetails.length + 4] = ''; // Suma Ancho
                   noSegmentsRow[cieloCommonDetails.length + 5] = ''; // Suma Alto
                   noSegmentsRow[cieloCommonDetails.length + 6] = ''; // Área Total Panel Cenefa

                    sheetData.push(noSegmentsRow);
            }
        } else if (item.type === 'cenefa') { // Add Cenefa details to Excel
            // Details for this Cenefa item (configuration)
            const cenefaCommonDetails = [
                getItemTypeName(item.type), // 0: Tipo Item
                item.number,                 // 1: Número Item
                '',                          // 2: Placeholder para Detalle/Dimensiones
                '',                          // 3: Nº Caras (empty for cenefa)
                '', '',                     // 4, 5: Panel Cara 1, Cara 2 (empty for cenefa)
                '',                          // 6: Tipo Panel Cielo (empty for cenefa)
                item.cenefaPanelType ? item.cenefaPanelType : '', // 7: Tipo Panel Cenefa
                '',                          // 8: Espaciamiento Postes (empty for cenefa)
                '',                          // 9: Pleno (empty for cenefa)
                '',                          // 10: Metros Descuento Angular (empty for cenefa)
                '',                           // 11: Estructura Doble (empty for cenefa)
                // New Cenefa Config Columns
                item.cenefaOrientation ? item.cenefaOrientation : '', // 12: Orientación Cenefa
                !isNaN(item.cenefaSides) ? item.cenefaSides : '', // 13: Lados/Caras Cenefa
                !isNaN(item.cenefaListonSpacing) ? item.cenefaListonSpacing.toFixed(2) : '', // 14: Esp. Listón Cenefa
                '',                          // 15: Esp. Soporte Cenefa (empty for cenefa)
                 // Total Columns
                '',                                                                                 // Suma Perímetros Segmentos (empty for Cenefa)
                '',                                                                                 // Ancho Total (muro) - Empty for Cenefa
                '',                                                                                 // Área Total (m²) del Cielo (empty for Cenefa)
                !isNaN(item.totalCenefaLargoSum) ? item.totalCenefaLargoSum.toFixed(2) : '', // Suma Largo Segmentos Cenefa
                !isNaN(item.totalCenefaAnchoSum) ? item.totalCenefaAnchoSum.toFixed(2) : '', // Suma Ancho Segmentos Cenefa
                !isNaN(item.totalCenefaAltoSum) ? item.totalCenefaAltoSum.toFixed(2) : '', // Suma Alto Segmentos Cenefa
                !isNaN(item.totalCenefaPanelArea) ? item.totalCenefaPanelArea.toFixed(2) : '' // Área Total Panel Cenefa (with sides)
            ];
             // Main row with item options and totals
            const cenefaSummaryRow = [...cenefaCommonDetails];
            cenefaSummaryRow[2] = 'Opciones:'; // Label in detail column
            sheetData.push(cenefaSummaryRow);

             // Row labeling the Segments section
             const cenefaSegmentsLabelRow = [...cenefaCommonDetails];
             cenefaSegmentsLabelRow[2] = 'Segmentos:'; // Label "Segmentos:"
              // Add empty cells for total columns
              // Need to fill the total columns at the end of the row
              cenefaSegmentsLabelRow.push('', '', '', '', '', '', ''); // Fill all total columns with empty strings
              sheetData.push(cenefaSegmentsLabelRow);


            if (item.segments && item.segments.length > 0) {
                  item.segments.forEach(seg => {
                      // Row for each individual Segment
                      const segmentRow = [...cenefaCommonDetails]; // Copy common details for this row
                      segmentRow[2] = `Seg ${seg.number}: ${seg.largo.toFixed(2)}m x ${seg.ancho.toFixed(2)}m x ${seg.alto.toFixed(2)}m`; // Segment dimensions
                       // Add empty cells for total columns (repeat empty totals for each segment row)
                       segmentRow.push('', '', '', '', '', '', ''); // Fill all total columns with empty strings
                       sheetData.push(segmentRow);
                  });
            } else {
                   // Row for "Sin segmentos válidos"
                   const noSegmentsRow = [...cenefaCommonDetails];
                   noSegmentsRow[2] = `- Sin segmentos válidos`; // Message
                    // Add empty cells for total columns
                    noSegmentsRow.push('', '', '', '', '', '', ''); // Fill all total columns with empty strings
                    sheetData.push(noSegmentsRow);
            }
        }
   });
   sheetData.push([]); // Fila en blanco para espaciar

   // Tabla de Totales de Materiales
    console.log("Añadiendo tabla de materiales totales al Excel.");
    sheetData.push(["Totales de Materiales (Cantidades a Comprar):"]);
   sheetData.push(["Material", "Cantidad", "Unidad"]);

   const sortedMaterials = Object.keys(lastCalculatedTotalMaterials).sort();
   sortedMaterials.forEach(material => {
       const cantidad = lastCalculatedTotalMaterials[material];
       const unidad = getMaterialUnit(material);
       // Usa el nombre del material directamente de la clave
        sheetData.push([material, cantidad, unidad]);
   });

   // Crea un libro y una hoja de Excel
   const wb = XLSX.utils.book_new();
   // `aoa_to_sheet` convierte un array de arrays (sheetData) a una hoja
   const ws = XLSX.utils.aoa_to_sheet(sheetData);

   // Opcional: Agregar estilos básicos (encabezados en negrita, etc.)
   // XLSX.js básico tiene limitaciones de estilo. Para estilos avanzados, se necesita otra librería.
   // Aquí no se agregan estilos complejos para mantener la simplicidad.

   // Agrega la hoja al libro
   XLSX.utils.book_append_sheet(wb, ws, "CalculoMateriales"); // "CalculoMateriales" es el nombre de la hoja

   // Genera y guarda el archivo Excel
   XLSX.writeFile(wb, `Calculo_Materiales_${new Date().toLocaleDateString('es-ES').replace(/\//g, '-')}.xlsx`); // Nombre del archivo con fecha
   console.log("Excel generado.");
};


// --- Event Listeners ---
// Asigna las funciones a los eventos de los botones
addItemBtn.addEventListener('click', createItemBlock); // Botón "Agregar Muro o Cielo"
calculateBtn.addEventListener('click', calculateMaterials); // Botón "Calcular Materiales"
generatePdfBtn.addEventListener('click', generatePDF); // Botón "Generar PDF"
generateExcelBtn.addEventListener('click', generateExcel); // Botón "Generar Excel"

// --- Configuración Inicial ---
// Agrega un ítem (Muro por defecto) al cargar la página
createItemBlock();
// Establece el estado inicial del botón de cálculo (habilitado si hay ítems)
toggleCalculateButtonState();


}); // Fin del evento DOMContentLoaded. Asegura que el script se ejecuta después de que la página esté completamente cargada.

