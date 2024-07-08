// ==UserScript==
// @name         Forms Office Question ID Appender Pro
// @author       Sebastian Cwynar (dvlart)
// @version      0.9.1
// @description  Appends question IDs to question text in forms.office.com with format options, inline copy buttons, and custom format generation. Features draggable and minimizable modal.
// @match        https://forms.office.com/*
// @downloadURL    https://github.com/dvlart/MSForms-ID/raw/main/Forms%20Office%20Question%20ID%20Appender%20Pro-0.9.user.js
// @updateURL    https://github.com/dvlart/MSForms-ID/raw/main/Forms%20Office%20Question%20ID%20Appender%20Pro-0.9.user.js
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    const BUTTON_ID = 'appendQuestionIdsButton';
    const MODAL_ID = 'optionsModal';
    const MINIMIZED_ID = 'minimizedModal';

    const styles = {
        mainButton: `
            position: fixed;
            top: 10px;
            right: 100px;
            z-index: 9999;
            padding: 5px 10px;
            background-color: #0078d4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        `,
        modal: `
            position: fixed;
            left: 50%;
            top: 50%;
            transform: translate(-50%, -50%);
            background-color: #BCFFF3;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2);
            z-index: 10000;
            width: 300px;
        `,
        modalHeader: `
            cursor: move;
            padding: 10px;
            background-color: #0078d4;
            color: white;
            margin: -20px -20px 10px -20px;
            border-radius: 8px 8px 0 0;
        `,
        closeButton: `
            float: right;
            background: none;
            border: none;
            color: white;
            font-size: 18px;
            cursor: pointer;
        `,
        minimizeButton: `
            float: right;
            background: none;
            border: none;
            color: white;
            font-size: 18px;
            cursor: pointer;
            margin-right: 10px;
        `,
        minimizedModal: `
            position: fixed;
            left: 10px;
            bottom: 10px;
            background-color: #0078d4;
            color: white;
            padding: 10px;
            border-radius: 4px;
            cursor: pointer;
            z-index: 10000;
        `,
        applyButton: `
            margin-top: 15px;
            padding: 5px 10px;
            background-color: #0078d4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        `,
        discardButton: `
            margin-top: 15px;
            padding: 5px 10px;
            background-color: #C70039;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        `,
        copyButton: `
            margin-left: 5px;
            padding: 2px 5px;
            background-color: #0078d4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
        `,
        customFormatInput: `
            width: 100%;
            padding: 5px;
            margin-top: 10px;
        `
    };

    function createMainButton() {
        const button = document.createElement('button');
        button.id = BUTTON_ID;
        button.textContent = 'Append Question IDs';
        button.style.cssText = styles.mainButton;
        button.addEventListener('click', showOptionsModal);
        return button;
    }

    function addMainButton() {
        if (document.getElementById(BUTTON_ID)) return;

        const button = createMainButton();
        const profileIcon = document.querySelector('button[aria-label="User account menu"]');
        if (profileIcon && profileIcon.parentNode) {
            profileIcon.parentNode.insertBefore(button, profileIcon);
        } else {
            document.body.appendChild(button);
        }
    }

    function createModalContent() {
        return `
            <div class="modal-header" style="${styles.modalHeader}">
                <span>Question ID Format Options</span>
                <button class="close-button" style="${styles.closeButton}">&times;</button>
                <button class="minimize-button" style="${styles.minimizeButton}">_</button>
            </div>
            <div>
                <input type="radio" id="format1" name="format" value="1" checked>
                <label for="format1">With @, {, and } (e.g., @{12345})</label>
            </div>
            <div>
                <input type="radio" id="format2" name="format" value="2">
                <label for="format2">Without @, {, and } (e.g., 12345)</label>
            </div>
            <div>
                <input type="radio" id="formatCustom" name="format" value="custom">
                <label for="formatCustom">Custom format</label>
            </div>
            <div>
                <input type="text" id="customFormatInput" placeholder="Enter custom text (e.g., Get response details)" style="${styles.customFormatInput}">
            </div>
            <div id="customFormatOptions" style="display: none; margin-top: 10px;">
                <div>
                    <input type="checkbox" id="includeAt" checked>
                    <label for="includeAt">Include @</label>
                </div>
                <div>
                    <input type="checkbox" id="includeBraces" checked>
                    <label for="includeBraces">Include { and }</label>
                </div>
            </div>
            <div style="margin-top: 10px;">
                <input type="checkbox" id="overrideText">
                <label for="overrideText">Override original text (show only ID)</label>
            </div>
            <div style="margin-top: 10px;">
                <input type="checkbox" id="addCopyButton" checked>
                <label for="addCopyButton">Add copy button next to IDs</label>
            </div>
            <button id="applyButton" style="${styles.applyButton}">Apply</button>
            <button id="discardButton" style="${styles.discardButton}">Close</button>
        `;
    }

    function makeDraggable(element) {
        let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
        element.querySelector('.modal-header').onmousedown = dragMouseDown;

        function dragMouseDown(e) {
            e = e || window.event;
            e.preventDefault();
            pos3 = e.clientX;
            pos4 = e.clientY;
            document.onmouseup = closeDragElement;
            document.onmousemove = elementDrag;
        }

        function elementDrag(e) {
            e = e || window.event;
            e.preventDefault();
            pos1 = pos3 - e.clientX;
            pos2 = pos4 - e.clientY;
            pos3 = e.clientX;
            pos4 = e.clientY;
            element.style.top = (element.offsetTop - pos2) + "px";
            element.style.left = (element.offsetLeft - pos1) + "px";
        }

        function closeDragElement() {
            document.onmouseup = null;
            document.onmousemove = null;
        }
    }

    function showOptionsModal() {
        const modal = document.createElement('div');
        modal.id = MODAL_ID;
        modal.style.cssText = styles.modal;
        modal.innerHTML = createModalContent();
        document.body.appendChild(modal);

        makeDraggable(modal);

        const formatCustom = document.getElementById('formatCustom');
        const customFormatOptions = document.getElementById('customFormatOptions');
        formatCustom.addEventListener('change', () => {
            customFormatOptions.style.display = formatCustom.checked ? 'block' : 'none';
        });

        document.getElementById('applyButton').addEventListener('click', () => {
            const options = {
                format: document.querySelector('input[name="format"]:checked').value,
                customText: document.getElementById('customFormatInput').value,
                includeAt: document.getElementById('includeAt').checked,
                includeBraces: document.getElementById('includeBraces').checked,
                override: document.getElementById('overrideText').checked,
                addCopyButton: document.getElementById('addCopyButton').checked
            };
            appendQuestionIds(options);
        });

        document.getElementById('discardButton').addEventListener('click', () => {
            document.body.removeChild(modal);
        });

        modal.querySelector('.close-button').addEventListener('click', () => {
            document.body.removeChild(modal);
        });

        modal.querySelector('.minimize-button').addEventListener('click', () => {
            minimizeModal(modal);
        });
    }

    function minimizeModal(modal) {
        document.body.removeChild(modal);
        const minimized = document.createElement('div');
        minimized.id = MINIMIZED_ID;
        minimized.style.cssText = styles.minimizedModal;
        minimized.textContent = 'Question ID Options';
        minimized.addEventListener('click', () => {
            document.body.removeChild(minimized);
            showOptionsModal();
        });
        document.body.appendChild(minimized);
    }

    function formatId(id, format, customText, includeAt, includeBraces) {
        if (format === "custom") {
            const formattedCustomText = customText.replace(/ /g, '_');
            let result = `outputs('${formattedCustomText}')?['body/${id}']`;
            if (includeBraces) {
                result = `{${result}}`;
            }
            if (includeAt) {
                result = `@${result}`;
            }
            return result;
        }
        return format === "1" ? `@{${id}}` : id;
    }

    function createCopyButton(formattedId) {
        const button = document.createElement('button');
        button.textContent = 'Copy';
        button.className = 'copyButton';
        button.style.cssText = styles.copyButton;
        button.addEventListener('click', (e) => {
            e.stopPropagation();
            navigator.clipboard.writeText(formattedId).then(() => {
                const originalText = button.textContent;
                button.textContent = 'Copied!';
                setTimeout(() => button.textContent = originalText, 1500);
            });
        });
        return button;
    }

    function appendQuestionIds(options) {
        const questions = document.querySelectorAll('div[data-automation-id="questionContent"]');
        let questionsModified = 0;

        questions.forEach(question => {
            const questionIdElement = question.querySelector('div[id*="QuestionId_"]');
            if (!questionIdElement) return;

            const questionText = question.querySelector('span[class="text-format-content "]');
            if (!questionText) return;

            const questionId = questionIdElement.getAttribute("id").split("_")[1];
            const formattedId = formatId(questionId, options.format, options.customText, options.includeAt, options.includeBraces);

            if (options.override) {
                questionText.textContent = formattedId;
            } else if (!questionText.textContent.includes(questionId)) {
                questionText.textContent += ` // ${formattedId}`;
            }

            if (options.addCopyButton && !question.querySelector('.copyButton')) {
                const copyButton = createCopyButton(formattedId);
                questionText.appendChild(copyButton);
            }

            questionsModified++;
        });

        //alert(`Operation complete. Modified ${questionsModified} question(s).`);
    }

    function init() {
        if (document.readyState === 'complete') {
            addMainButton();
        } else {
            setTimeout(init, 1000);
        }
    }

    init();
})();
