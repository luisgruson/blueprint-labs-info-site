// Excel Card Interface
class ExcelCardInterface {
    constructor() {
        this.selectedCell = null;
        this.cards = {};
        this.currentSheet = 'Automate workflows';
        this.sheets = {
            'Automate workflows': { cards: {} },
            'Web scraping services': { cards: {} },
            'CRM management': { cards: {} },
            'Property Assessment': { cards: {} },
            'Dynamic Modeling': { cards: {} },
            'Lead Generation': { cards: {} }
        };
        this.sheetNames = Object.keys(this.sheets);
        this.currentSheetIndex = 0;
        this.autoCycleInterval = null;
        this.userInteracted = false;
        this.init();
        this.populateWithSampleCards();
        this.startAutoCycle();
    }

    init() {
        const spreadsheet = document.getElementById('spreadsheet');
        
        // Create a simple container for the sheet content
        const sheetContainer = document.createElement('div');
        sheetContainer.className = 'sheet-container';
        sheetContainer.style.display = 'flex';
        sheetContainer.style.justifyContent = 'center';
        sheetContainer.style.alignItems = 'center';
        sheetContainer.style.height = '100%';
        sheetContainer.style.width = '100%';
        sheetContainer.style.padding = '20px';
        sheetContainer.style.boxSizing = 'border-box';
        
        // Create h3 element for the sheet title
        const titleElement = document.createElement('h3');
        titleElement.className = 'sheet-title';
        titleElement.style.fontFamily = "'HomeVideo', Arial, sans-serif";
        titleElement.style.fontSize = '2rem';
        titleElement.style.color = 'black';
        titleElement.style.textAlign = 'center';
        titleElement.style.margin = '0';
        titleElement.style.width = '100%';
        titleElement.style.whiteSpace = 'nowrap';
        titleElement.style.overflow = 'visible';
        
        sheetContainer.appendChild(titleElement);
        spreadsheet.appendChild(sheetContainer);
        
        // Sheet tab functionality
        this.setupSheetTabs();
    }

    selectCell(cell) {
        if (this.selectedCell) {
            this.selectedCell.classList.remove('selected');
        }
        
        cell.classList.add('selected');
        this.selectedCell = cell;
        
        // Update name box
        document.querySelector('.name-box').value = cell.dataset.address;
        
        // Update formula bar
        const formulaInput = document.querySelector('.formula-input');
        const existingCard = this.cards[cell.dataset.address];
        formulaInput.value = existingCard ? existingCard.content : '';
        formulaInput.focus();
    }

    addCard(address, content) {
        if (!content.trim()) return;

        const cell = document.querySelector(`[data-address="${address}"]`);
        if (!cell) return;

        // Remove existing card if any
        if (this.cards[address]) {
            const existingCard = cell.querySelector('.card');
            if (existingCard) {
                existingCard.remove();
            }
        }

        // Create new card
        const card = document.createElement('div');
        card.className = 'card';
        card.textContent = content;
        card.title = content; // Tooltip for longer text
        
        // Card editing disabled
        // card.addEventListener('click', (e) => {
        //     e.stopPropagation();
        //     this.editCard(address);
        // });

        cell.appendChild(card);
        this.cards[address] = { content, element: card };
    }

    editCard(address) {
        const formulaInput = document.querySelector('.formula-input');
        formulaInput.value = this.cards[address].content;
        formulaInput.focus();
        formulaInput.select();
    }

    populateWithSampleCards() {
        // Display the current sheet title
        this.displaySheetTitle();
    }

    // API for external use
    addCardToCell(address, content) {
        this.addCard(address, content);
    }

    removeCard(address) {
        const cell = document.querySelector(`[data-address="${address}"]`);
        if (cell && this.cards[address]) {
            const card = cell.querySelector('.card');
            if (card) card.remove();
            delete this.cards[address];
        }
    }

    getAllCards() {
        return { ...this.cards };
    }

    clearAllCards() {
        Object.keys(this.cards).forEach(address => {
            this.removeCard(address);
        });
    }

    setupSheetTabs() {
        const sheetTabs = document.querySelectorAll('.sheet-tab');
        console.log('Found sheet tabs:', sheetTabs.length); // Debug log
        
        sheetTabs.forEach(tab => {
            tab.addEventListener('click', (e) => {
                e.preventDefault();
                const sheetName = e.target.textContent.trim();
                console.log('Tab clicked:', sheetName); // Debug log
                this.stopAutoCycle();
                this.userInteracted = true;
                this.switchToSheet(sheetName);
            });
        });
    }

    switchToSheet(sheetName) {
        console.log('Switching to sheet:', sheetName); // Debug log
        
        // Update active tab
        document.querySelectorAll('.sheet-tab').forEach(tab => {
            tab.classList.remove('active');
        });
        
        // Find and activate the correct tab
        const tabs = document.querySelectorAll('.sheet-tab');
        tabs.forEach(tab => {
            if (tab.textContent.trim() === sheetName) {
                tab.classList.add('active');
            }
        });

        // Save current sheet data
        this.sheets[this.currentSheet].cards = { ...this.cards };

        // Clear current spreadsheet
        this.clearAllCards();

        // Switch to new sheet
        this.currentSheet = sheetName;
        this.cards = { ...this.sheets[sheetName].cards };

        // Rebuild the spreadsheet with new sheet data
        this.rebuildSpreadsheet();
    }

    rebuildSpreadsheet() {
        // Display the current sheet title
        this.displaySheetTitle();
    }

    startAutoCycle() {
        if (this.autoCycleInterval) {
            clearInterval(this.autoCycleInterval);
        }
        
        this.autoCycleInterval = setInterval(() => {
            if (!this.userInteracted) {
                this.cycleToNextSheet();
            }
        }, 6000); // 6 seconds
    }

    stopAutoCycle() {
        if (this.autoCycleInterval) {
            clearInterval(this.autoCycleInterval);
            this.autoCycleInterval = null;
        }
    }

    cycleToNextSheet() {
        this.currentSheetIndex = (this.currentSheetIndex + 1) % this.sheetNames.length;
        const nextSheetName = this.sheetNames[this.currentSheetIndex];
        console.log('Auto-cycling to:', nextSheetName);
        this.switchToSheet(nextSheetName);
    }

    displaySheetTitle() {
        const titleElement = document.querySelector('.sheet-title');
        if (titleElement) {
            // Clear the current content
            titleElement.innerHTML = '';
            
            // Create the service card content
            const serviceData = this.getServiceData(this.currentSheet);
            
            // Create h4 title
            const title = document.createElement('h4');
            title.style.fontFamily = "'MS Sans Serif', sans-serif";
            title.style.fontSize = '14pt';
            title.style.fontWeight = 'bold';
            title.style.color = 'black';
            title.style.margin = '0 0 10px 0';
            title.style.textAlign = 'center';
            title.textContent = serviceData.title;
            
            // Create h5 description
            const description = document.createElement('h5');
            description.style.fontFamily = "'MS Sans Serif', sans-serif";
            description.style.fontSize = '10pt';
            description.style.fontWeight = 'normal';
            description.style.color = '#333';
            description.style.margin = '0 0 15px 0';
            description.style.textAlign = 'center';
            description.style.lineHeight = '1.4';
            description.textContent = serviceData.description;
            
            // Create learn more button
            const button = document.createElement('button');
            button.textContent = 'Learn More';
            button.style.fontFamily = "'MS Sans Serif', sans-serif";
            button.style.fontSize = '9pt';
            button.style.backgroundColor = '#0054e3';
            button.style.color = 'white';
            button.style.border = '1px outset #0054e3';
            button.style.padding = '6px 16px';
            button.style.cursor = 'pointer';
            button.style.borderRadius = '2px';
            button.style.margin = '0 auto';
            button.style.display = 'block';
            
            // Add hover effect
            button.addEventListener('mouseenter', () => {
                button.style.backgroundColor = '#003d99';
            });
            button.addEventListener('mouseleave', () => {
                button.style.backgroundColor = '#0054e3';
            });
            
            // Append all elements
            titleElement.appendChild(title);
            titleElement.appendChild(description);
            titleElement.appendChild(button);
        }
    }
    
    getServiceData(serviceName) {
        const serviceData = {
            'Automate workflows': {
                title: 'Automate Workflows',
                description: 'Streamline your business processes with intelligent automation solutions that reduce manual work and increase efficiency.'
            },
            'Web scraping services': {
                title: 'Web Scraping Services',
                description: 'Extract valuable data from websites to gain competitive insights and market intelligence for your real estate business.'
            },
            'CRM management': {
                title: 'CRM Management',
                description: 'Optimize customer relationships with AI-powered CRM solutions that help you manage leads and close more deals.'
            },
            'Property Assessment': {
                title: 'Property Assessment',
                description: 'Accurate property valuations using advanced AI algorithms that analyze market data and property characteristics.'
            },
            'Dynamic Modeling': {
                title: 'Dynamic Modeling',
                description: 'Create predictive models for market trends and property values to make data-driven investment decisions.'
            },
            'Lead Generation': {
                title: 'Lead Generation',
                description: 'Generate high-quality leads through intelligent targeting systems that identify potential clients and opportunities.'
            }
        };
        
        return serviceData[serviceName] || { title: serviceName, description: 'Service description coming soon.' };
    }
}

// Initialize the interface
const excelInterface = new ExcelCardInterface();

// Expose API globally for integration
window.ExcelCardInterface = {
    addCard: (address, content) => excelInterface.addCardToCell(address, content),
    removeCard: (address) => excelInterface.removeCard(address),
    getAllCards: () => excelInterface.getAllCards(),
    clearAll: () => excelInterface.clearAllCards()
};
