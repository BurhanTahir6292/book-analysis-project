let booksData = [];
let chartInstances = {};
let activeTab = 'visualizations';
let viewMode = 'grid'; // 'grid' or 'list'
let selectedBookId = null;

// Simple Sentiment & Emotion analyzer for local additions/edits
function estimateSentiment(text, rating) {
    if (!text) {
        if (rating >= 4) return { score: 0.6, class: 'Positive', emotion: 'Joy' };
        if (rating <= 2) return { score: -0.4, class: 'Negative', emotion: rating === 2 ? 'Disappointment' : 'Anger' };
        return { score: 0.0, class: 'Neutral', emotion: 'Neutral' };
    }
    const positiveWords = ['great', 'good', 'masterpiece', 'excellent', 'amazing', 'fantastic', 'love', 'liked', 'wonderful', 'beautiful', 'outstanding', 'worth', 'superb', 'enjoyed', 'recommended'];
    const negativeWords = ['bad', 'awful', 'terrible', 'worst', 'poor', 'disappointing', 'hate', 'boring', 'nonsensical', 'struggled', 'waste', 'dull', 'slow', 'horrible', 'avoid'];
    
    let score = 0;
    const tokens = text.toLowerCase().split(/\W+/);
    tokens.forEach(word => {
        if (positiveWords.includes(word)) score += 0.25;
        if (negativeWords.includes(word)) score -= 0.25;
    });

    // Blend with rating
    if (rating >= 4) score += 0.2;
    if (rating <= 2) score -= 0.2;

    // Clamp score
    score = Math.max(-1, Math.min(1, score));

    let sentimentClass = 'Neutral';
    if (score > 0.05) sentimentClass = 'Positive';
    else if (score < -0.05) sentimentClass = 'Negative';

    // Emotion Detection Logic mirroring Python
    const text_lower = text.toLowerCase();
    const joy_words = ['masterpiece', 'recommend', 'great', 'enjoyed', 'liked', 'fantastic', 'best', 'solid', 'entertaining', 'good', 'wonderful', 'loved', 'love', 'superb'];
    const surprise_words = ['amazing', 'incredible', 'wonder', 'surprised'];
    const disappointment_words = ['disappointing', 'disappointment', 'expected a bit more', 'below average', 'letdown'];
    const sadness_words = ['dull', 'slow', 'dragged out', 'struggled', 'boring', 'sad'];
    const anger_words = ['terrible', 'awful', 'worst', 'waste', 'hated', 'nonsensical', 'zero redeeming', 'poorly written'];

    const countOccurrences = (str, word) => {
        let count = 0;
        let pos = str.indexOf(word);
        while (pos !== -1) {
            count++;
            pos = str.indexOf(word, pos + word.length);
        }
        return count;
    };

    let score_joy = joy_words.reduce((sum, w) => sum + countOccurrences(text_lower, w), 0);
    let score_surprise = surprise_words.reduce((sum, w) => sum + countOccurrences(text_lower, w), 0);
    let score_disappointment = disappointment_words.reduce((sum, w) => sum + countOccurrences(text_lower, w), 0);
    let score_sadness = sadness_words.reduce((sum, w) => sum + countOccurrences(text_lower, w), 0);
    let score_anger = anger_words.reduce((sum, w) => sum + countOccurrences(text_lower, w), 0);

    if (rating === 5) {
        score_joy += 2;
        score_surprise += 1;
    } else if (rating === 4) {
        score_joy += 1;
    } else if (rating === 2) {
        score_disappointment += 1;
        score_sadness += 1;
    } else if (rating === 1) {
        score_anger += 2;
        score_disappointment += 1;
    }

    const scores = {
        'Joy': score_joy,
        'Surprise': score_surprise,
        'Disappointment': score_disappointment,
        'Sadness': score_sadness,
        'Anger': score_anger
    };

    let max_emotion = 'Neutral';
    let max_val = 0;
    for (const [emotion, val] of Object.entries(scores)) {
        if (val > max_val) {
            max_val = val;
            max_emotion = emotion;
        }
    }

    return { 
        score: parseFloat(score.toFixed(3)), 
        class: sentimentClass,
        emotion: max_emotion
    };
}

// Initial Data Load
document.addEventListener('DOMContentLoaded', () => {
    fetch('dashboard_data.json')
        .then(response => response.json())
        .then(data => {
            booksData = data;
            initApp();
        })
        .catch(error => {
            console.error('Error loading dashboard data:', error);
            showToast('❌ Error loading dashboard data. Check console.');
        });
});

// App Initialization
function initApp() {
    setupTabNavigation();
    setupViewToggles();
    setupSearchFilters();
    setupModals();
    setupForms();
    
    renderDashboard();
}

// Global Render (called when data changes)
function renderDashboard() {
    renderKPIs(booksData);
    renderVisualizations(booksData);
    renderLibrary(booksData);
    renderAnalytics(booksData);
}

// Tab Navigation logic
function setupTabNavigation() {
    const tabs = ['vis', 'books', 'analytics'];
    tabs.forEach(tabId => {
        const btn = document.getElementById(`tab-${tabId}`);
        if (btn) {
            btn.addEventListener('click', () => {
                // Remove active classes
                document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                
                // Add active classes
                btn.classList.add('active');
                const contentId = `tab-${btn.dataset.tab}`;
                document.getElementById(contentId).classList.add('active');
                activeTab = btn.dataset.tab;
                
                // Trigger chart resizing if active tab is visualizations
                if (activeTab === 'visualizations') {
                    Object.values(chartInstances).forEach(chart => chart.resize());
                }
            });
        }
    });
}

// View Toggles logic
function setupViewToggles() {
    const gridBtn = document.getElementById('view-grid');
    const listBtn = document.getElementById('view-list');
    
    gridBtn.addEventListener('click', () => {
        gridBtn.classList.add('active');
        listBtn.classList.remove('active');
        viewMode = 'grid';
        document.getElementById('book-grid').style.display = 'grid';
        document.getElementById('book-list-view').style.display = 'none';
    });

    listBtn.addEventListener('click', () => {
        listBtn.classList.add('active');
        gridBtn.classList.remove('active');
        viewMode = 'list';
        document.getElementById('book-grid').style.display = 'none';
        document.getElementById('book-list-view').style.display = 'block';
    });
}

// Setup Search & Filters
function setupSearchFilters() {
    const inputs = ['search-input', 'filter-rating', 'filter-sentiment', 'sort-by'];
    inputs.forEach(id => {
        document.getElementById(id).addEventListener('input', () => {
            renderLibrary(booksData);
        });
    });
}

// KPI Calculation
function renderKPIs(data) {
    const totalBooks = data.length;
    const kpiContainer = document.getElementById('kpi-container');
    
    if (totalBooks === 0) {
        kpiContainer.innerHTML = `
            <div class="kpi-card"><div class="kpi-value">0</div><div class="kpi-label">Books Analyzed</div></div>
            <div class="kpi-card"><div class="kpi-value">£0.00</div><div class="kpi-label">Average Price</div></div>
            <div class="kpi-card"><div class="kpi-value">0%</div><div class="kpi-label">Positive Sentiment</div></div>
        `;
        return;
    }
    
    const avgPrice = data.reduce((sum, book) => sum + parseFloat(book.Price), 0) / totalBooks;
    
    let positiveCount = 0;
    data.forEach(book => {
        if (book.Sentiment_Class === 'Positive') positiveCount++;
    });
    
    const positivePercentage = Math.round((positiveCount / totalBooks) * 100);

    kpiContainer.innerHTML = `
        <div class="kpi-card">
            <div class="kpi-value">${totalBooks}</div>
            <div class="kpi-label">Books Analyzed</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-value">£${avgPrice.toFixed(2)}</div>
            <div class="kpi-label">Average Price</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-value" style="background: linear-gradient(135deg, #34d399, #10b981); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">
                ${positivePercentage}%
            </div>
            <div class="kpi-label">Positive Sentiment</div>
        </div>
    `;
}

// Interactive Visualizations Rendering
function renderVisualizations(data) {
    Chart.defaults.color = '#94a3b8';
    Chart.defaults.font.family = "'Inter', sans-serif";
    
    // Destroy previous charts
    Object.keys(chartInstances).forEach(key => {
        if (chartInstances[key]) {
            chartInstances[key].destroy();
        }
    });

    if (data.length === 0) return;

    // 1. Price Distribution Chart
    const prices = data.map(d => parseFloat(d.Price));
    const minPrice = Math.floor(Math.min(...prices)) || 0;
    const maxPrice = Math.ceil(Math.max(...prices)) || 100;
    const bins = 8;
    const binSize = (maxPrice - minPrice) / bins;
    const priceBins = Array(bins).fill(0);
    const priceLabels = Array(bins).fill('');
    
    for (let i = 0; i < bins; i++) {
        priceLabels[i] = `£${(minPrice + i * binSize).toFixed(0)}-£${(minPrice + (i + 1) * binSize).toFixed(0)}`;
    }
    prices.forEach(p => {
        let binIdx = Math.floor((p - minPrice) / binSize);
        if (binIdx >= bins) binIdx = bins - 1;
        if (binIdx < 0) binIdx = 0;
        priceBins[binIdx]++;
    });

    const ctxPrice = document.getElementById('priceDistChart').getContext('2d');
    chartInstances.price = new Chart(ctxPrice, {
        type: 'bar',
        data: {
            labels: priceLabels,
            datasets: [{
                label: 'Number of Books',
                data: priceBins,
                backgroundColor: 'rgba(59, 130, 246, 0.65)',
                borderColor: '#3b82f6',
                borderWidth: 1.5,
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { backgroundColor: 'rgba(15, 23, 42, 0.95)', padding: 12 }
            },
            scales: {
                y: { grid: { color: 'rgba(255, 255, 255, 0.05)' }, beginAtZero: true },
                x: { grid: { display: false } }
            }
        }
    });

    // Update Price Stats on Chart
    const avgPrice = prices.reduce((a,b)=>a+b, 0)/prices.length;
    document.getElementById('price-stats').innerHTML = `
        <span>Min: <strong>£${minPrice.toFixed(2)}</strong></span>
        <span>Avg: <strong>£${avgPrice.toFixed(2)}</strong></span>
        <span>Max: <strong>£${maxPrice.toFixed(2)}</strong></span>
    `;

    // 2. Average Price by Rating
    const ratingSums = {1:0, 2:0, 3:0, 4:0, 5:0};
    const ratingCounts = {1:0, 2:0, 3:0, 4:0, 5:0};
    data.forEach(d => {
        const rating = d['Numeric Rating'];
        if (ratingSums[rating] !== undefined) {
            ratingSums[rating] += parseFloat(d.Price);
            ratingCounts[rating]++;
        }
    });
    const avgPriceByRating = [1, 2, 3, 4, 5].map(r => ratingCounts[r] > 0 ? ratingSums[r] / ratingCounts[r] : 0);

    const ctxAvg = document.getElementById('avgPriceChart').getContext('2d');
    chartInstances.avgPrice = new Chart(ctxAvg, {
        type: 'bar',
        data: {
            labels: ['1 Star', '2 Stars', '3 Stars', '4 Stars', '5 Stars'],
            datasets: [{
                label: 'Avg Price (£)',
                data: avgPriceByRating,
                backgroundColor: 'rgba(139, 92, 246, 0.65)',
                borderColor: '#8b5cf6',
                borderWidth: 1.5,
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { backgroundColor: 'rgba(15, 23, 42, 0.95)', padding: 12 }
            },
            scales: {
                y: { grid: { color: 'rgba(255, 255, 255, 0.05)' }, beginAtZero: true },
                x: { grid: { display: false } }
            }
        }
    });

    // 3. Sentiment Doughnut
    const sentimentCounts = { 'Positive': 0, 'Neutral': 0, 'Negative': 0 };
    data.forEach(d => {
        if (sentimentCounts[d.Sentiment_Class] !== undefined) {
            sentimentCounts[d.Sentiment_Class]++;
        }
    });

    const ctxSentiment = document.getElementById('sentimentChart').getContext('2d');
    chartInstances.sentiment = new Chart(ctxSentiment, {
        type: 'doughnut',
        data: {
            labels: ['Positive', 'Neutral', 'Negative'],
            datasets: [{
                data: [sentimentCounts['Positive'], sentimentCounts['Neutral'], sentimentCounts['Negative']],
                backgroundColor: [
                    'rgba(16, 185, 129, 0.75)',
                    'rgba(100, 116, 139, 0.75)',
                    'rgba(239, 68, 68, 0.75)'
                ],
                borderColor: '#1e293b',
                borderWidth: 2,
                hoverOffset: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '70%',
            plugins: {
                legend: { position: 'bottom', labels: { padding: 15 } },
                tooltip: { backgroundColor: 'rgba(15, 23, 42, 0.95)', padding: 12 }
            }
        }
    });

    // Update Sentiment Stats on chart Card
    document.getElementById('sentiment-stats').innerHTML = `
        <span style="color:#10b981">Pos: <strong>${sentimentCounts['Positive']}</strong></span>
        <span style="color:#94a3b8">Neu: <strong>${sentimentCounts['Neutral']}</strong></span>
        <span style="color:#ef4444">Neg: <strong>${sentimentCounts['Negative']}</strong></span>
    `;

    // 4. Genre Chart (Horizontal Bar Chart)
    const genresMap = {};
    data.forEach(d => {
        const genre = d.Genre || 'Unknown';
        genresMap[genre] = (genresMap[genre] || 0) + 1;
    });
    // Sort genres by count
    const sortedGenres = Object.entries(genresMap).sort((a,b)=>b[1]-a[1]).slice(0, 7);
    const genreLabels = sortedGenres.map(g => g[0]);
    const genreCounts = sortedGenres.map(g => g[1]);

    const ctxGenre = document.getElementById('genreChart').getContext('2d');
    chartInstances.genre = new Chart(ctxGenre, {
        type: 'bar',
        data: {
            labels: genreLabels,
            datasets: [{
                label: 'Books',
                data: genreCounts,
                backgroundColor: 'rgba(16, 185, 129, 0.65)',
                borderColor: '#10b981',
                borderWidth: 1.5,
                borderRadius: 6
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { backgroundColor: 'rgba(15, 23, 42, 0.95)' }
            },
            scales: {
                x: { grid: { color: 'rgba(255, 255, 255, 0.05)' }, beginAtZero: true, ticks: { stepSize: 1 } },
                y: { grid: { display: false } }
            }
        }
    });

    // 5. Rating Distribution Bar Chart
    const ratingDistCounts = {1:0, 2:0, 3:0, 4:0, 5:0};
    data.forEach(d => {
        const rating = d['Numeric Rating'];
        if (ratingDistCounts[rating] !== undefined) ratingDistCounts[rating]++;
    });

    const ctxRatingDist = document.getElementById('ratingDistChart').getContext('2d');
    chartInstances.ratingDist = new Chart(ctxRatingDist, {
        type: 'bar',
        data: {
            labels: ['★☆☆☆☆', '★★☆☆☆', '★★★☆☆', '★★★★☆', '★★★★★'],
            datasets: [{
                label: 'Books',
                data: [ratingDistCounts[1], ratingDistCounts[2], ratingDistCounts[3], ratingDistCounts[4], ratingDistCounts[5]],
                backgroundColor: 'rgba(245, 158, 11, 0.65)',
                borderColor: '#f59e0b',
                borderWidth: 1.5,
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { backgroundColor: 'rgba(15, 23, 42, 0.95)' }
            },
            scales: {
                y: { grid: { color: 'rgba(255, 255, 255, 0.05)' }, beginAtZero: true, ticks: { stepSize: 1 } },
                x: { grid: { display: false } }
            }
        }
    });

    // 6. Sentiment Polarity Scatter Chart
    const scatterData = data.map(d => ({
        x: parseFloat(d.Price),
        y: parseFloat(d.Sentiment_Polarity) || 0,
        title: d.Title
    }));

    const ctxScatter = document.getElementById('scatterChart').getContext('2d');
    chartInstances.scatter = new Chart(ctxScatter, {
        type: 'scatter',
        data: {
            datasets: [{
                label: 'Books',
                data: scatterData,
                backgroundColor: 'rgba(167, 139, 250, 0.8)',
                borderColor: '#a78bfa',
                borderWidth: 1,
                pointRadius: 6,
                pointHoverRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(15, 23, 42, 0.95)',
                    callbacks: {
                        label: function(context) {
                            const item = context.raw;
                            return `${item.title}: Price: £${item.x.toFixed(2)}, Polarity: ${item.y}`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    title: { display: true, text: 'Price (£)', color: '#94a3b8' },
                    grid: { color: 'rgba(255, 255, 255, 0.05)' }
                },
                y: {
                    title: { display: true, text: 'Sentiment Polarity (-1 to 1)', color: '#94a3b8' },
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    min: -1,
                    max: 1
                }
            }
        }
    });

    // 7. Decade Publication Timeline Chart
    const decadeCounts = {};
    data.forEach(d => {
        const year = parseInt(d.PublishYear);
        if (!isNaN(year) && year > 0) {
            const decade = Math.floor(year / 10) * 10;
            decadeCounts[decade] = (decadeCounts[decade] || 0) + 1;
        } else if (year < 0) {
            decadeCounts['BCE'] = (decadeCounts['BCE'] || 0) + 1;
        }
    });

    // Sort Decades
    const sortedDecades = Object.entries(decadeCounts).sort((a,b) => {
        if (a[0] === 'BCE') return -1;
        if (b[0] === 'BCE') return 1;
        return parseInt(a[0]) - parseInt(b[0]);
    });

    const decadeLabels = sortedDecades.map(item => item[0] === 'BCE' ? 'Classical Antiquity' : `${item[0]}s`);
    const decadeData = sortedDecades.map(item => item[1]);

    const ctxDecade = document.getElementById('decadeChart').getContext('2d');
    chartInstances.decade = new Chart(ctxDecade, {
        type: 'line',
        data: {
            labels: decadeLabels,
            datasets: [{
                label: 'Books Published',
                data: decadeData,
                fill: true,
                backgroundColor: 'rgba(59, 130, 246, 0.15)',
                borderColor: '#3b82f6',
                borderWidth: 2.5,
                tension: 0.35,
                pointRadius: 5,
                pointBackgroundColor: '#3b82f6'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { backgroundColor: 'rgba(15, 23, 42, 0.95)' }
            },
            scales: {
                y: { grid: { color: 'rgba(255, 255, 255, 0.05)' }, beginAtZero: true, ticks: { stepSize: 1 } },
                x: { grid: { color: 'rgba(255, 255, 255, 0.03)' } }
            }
        }
    });

    // 8. Emotion Distribution Chart (Radar)
    const emotionCounts = { 'Joy': 0, 'Surprise': 0, 'Disappointment': 0, 'Sadness': 0, 'Anger': 0, 'Neutral': 0 };
    data.forEach(d => {
        const emo = d.Emotion || 'Neutral';
        if (emotionCounts[emo] !== undefined) {
            emotionCounts[emo]++;
        }
    });

    const ctxEmotion = document.getElementById('emotionChart').getContext('2d');
    chartInstances.emotion = new Chart(ctxEmotion, {
        type: 'radar',
        data: {
            labels: ['Joy', 'Surprise', 'Disappointment', 'Sadness', 'Anger', 'Neutral'],
            datasets: [{
                label: 'Number of Books',
                data: [
                    emotionCounts['Joy'],
                    emotionCounts['Surprise'],
                    emotionCounts['Disappointment'],
                    emotionCounts['Sadness'],
                    emotionCounts['Anger'],
                    emotionCounts['Neutral']
                ],
                backgroundColor: 'rgba(139, 92, 246, 0.2)',
                borderColor: '#8b5cf6',
                pointBackgroundColor: '#8b5cf6',
                pointBorderColor: '#fff',
                pointHoverBackgroundColor: '#fff',
                pointHoverBorderColor: '#8b5cf6',
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: { backgroundColor: 'rgba(15, 23, 42, 0.95)', padding: 12 }
            },
            scales: {
                r: {
                    angleLines: { color: 'rgba(255, 255, 255, 0.08)' },
                    grid: { color: 'rgba(255, 255, 255, 0.08)' },
                    pointLabels: { color: '#94a3b8', font: { size: 11, weight: 'bold' } },
                    ticks: { display: false },
                    suggestedMin: 0
                }
            }
        }
    });

    // Update Emotion Stats on chart Card
    document.getElementById('emotion-stats').innerHTML = `
        <span style="color:#10b981">Joy: <strong>${emotionCounts['Joy']}</strong></span>
        <span style="color:#a78bfa">Surprise: <strong>${emotionCounts['Surprise']}</strong></span>
        <span style="color:#fbbf24">Disappointment: <strong>${emotionCounts['Disappointment']}</strong></span>
        <span style="color:#60a5fa">Sadness: <strong>${emotionCounts['Sadness']}</strong></span>
        <span style="color:#ef4444">Anger: <strong>${emotionCounts['Anger']}</strong></span>
        <span style="color:#94a3b8">Neutral: <strong>${emotionCounts['Neutral']}</strong></span>
    `;
}

// Render Books Library
function renderLibrary(data) {
    const search = document.getElementById('search-input').value.toLowerCase();
    const ratingFilter = document.getElementById('filter-rating').value;
    const sentimentFilter = document.getElementById('filter-sentiment').value;
    const sortBy = document.getElementById('sort-by').value;

    // Filter data
    let filtered = data.filter(book => {
        const matchesSearch = book.Title.toLowerCase().includes(search) || 
                              book.Author.toLowerCase().includes(search) || 
                              (book.Genre && book.Genre.toLowerCase().includes(search));
        
        const matchesRating = ratingFilter === "" || book['Numeric Rating'] == ratingFilter;
        const matchesSentiment = sentimentFilter === "" || book.Sentiment_Class === sentimentFilter;

        return matchesSearch && matchesRating && matchesSentiment;
    });

    // Sort data
    filtered.sort((a, b) => {
        if (sortBy === 'title') return a.Title.localeCompare(b.Title);
        if (sortBy === 'price-asc') return parseFloat(a.Price) - parseFloat(b.Price);
        if (sortBy === 'price-desc') return parseFloat(b.Price) - parseFloat(a.Price);
        if (sortBy === 'rating-desc') return b['Numeric Rating'] - a['Numeric Rating'];
        if (sortBy === 'year-desc') return (b.PublishYear || 0) - (a.PublishYear || 0);
        if (sortBy === 'year-asc') return (a.PublishYear || 0) - (b.PublishYear || 0);
        return 0;
    });

    // Update Counter
    document.getElementById('library-count').textContent = `Showing ${filtered.length} of ${data.length} books`;

    // Render Grid View
    const gridContainer = document.getElementById('book-grid');
    gridContainer.innerHTML = '';

    if (filtered.length === 0) {
        gridContainer.innerHTML = '<div style="grid-column: 1/-1; text-align: center; padding: 3rem; color: var(--text-secondary);">No books match your filters.</div>';
    }

    filtered.forEach(book => {
        const card = document.createElement('div');
        card.className = 'book-card glass-panel';
        card.setAttribute('data-id', book.id);
        
        const coverColor = book.CoverColor || '#4b5563';
        const stars = '★'.repeat(book['Numeric Rating']) + '☆'.repeat(5 - book['Numeric Rating']);

        card.innerHTML = `
            <div class="card-quick-actions">
                <button class="quick-btn quick-edit" title="Edit" data-id="${book.id}">✏️</button>
                <button class="quick-btn quick-delete" title="Delete" data-id="${book.id}">🗑️</button>
            </div>
            <div class="book-cover-container">
                <div class="book-mock-cover" style="background: linear-gradient(135deg, ${coverColor}, #1e293b);">
                    <div class="book-spine-line"></div>
                    <div class="cover-title">${book.Title}</div>
                    <div class="cover-author">${book.Author}</div>
                </div>
            </div>
            <div class="book-details">
                <div>
                    <h4 class="book-title" title="${book.Title}">${book.Title}</h4>
                    <p class="book-author">${book.Author}</p>
                </div>
                <div class="book-card-middle">
                    <div class="book-card-rating">${stars}</div>
                    <div class="book-card-price">£${parseFloat(book.Price).toFixed(2)}</div>
                </div>
                <div class="book-card-footer">
                    <div class="book-genre-badge">${book.Genre || 'Unknown'}</div>
                    <div style="display: flex; gap: 0.3rem;">
                        <span class="badge ${book.Sentiment_Class.toLowerCase()}">${book.Sentiment_Class}</span>
                        <span class="badge ${(book.Emotion || 'Neutral').toLowerCase()}">${book.Emotion || 'Neutral'}</span>
                    </div>
                </div>
            </div>
        `;

        // Card Click opens preview modal
        card.addEventListener('click', (e) => {
            if (e.target.closest('.quick-btn')) return; // Ignore if clicked action buttons
            openBookDetailModal(book.id);
        });

        // Quick Edit click
        card.querySelector('.quick-edit').addEventListener('click', (e) => {
            e.stopPropagation();
            openEditModal(book.id);
        });

        // Quick Delete click
        card.querySelector('.quick-delete').addEventListener('click', (e) => {
            e.stopPropagation();
            openConfirmDeleteModal(book.id);
        });

        gridContainer.appendChild(card);
    });

    // Render Table View
    const tbody = document.getElementById('books-tbody');
    tbody.innerHTML = '';

    filtered.forEach(book => {
        const row = document.createElement('tr');
        row.setAttribute('data-id', book.id);
        
        const stars = '★'.repeat(book['Numeric Rating']) + '☆'.repeat(5 - book['Numeric Rating']);
        
        row.innerHTML = `
            <td><strong>${book.Title}</strong></td>
            <td>${book.Author}</td>
            <td>${book.Genre || 'Unknown'}</td>
            <td>${book.PublishYear || 'N/A'}</td>
            <td style="color: #fbbf24; letter-spacing: 1px; white-space: nowrap;">${stars}</td>
            <td><strong>£${parseFloat(book.Price).toFixed(2)}</strong></td>
            <td><span class="badge ${book.Sentiment_Class.toLowerCase()}">${book.Sentiment_Class}</span></td>
            <td><span class="badge ${(book.Emotion || 'Neutral').toLowerCase()}">${book.Emotion || 'Neutral'}</span></td>
            <td>
                <div class="table-action-btns">
                    <button class="quick-btn quick-edit" title="Edit" data-id="${book.id}">✏️</button>
                    <button class="quick-btn quick-delete" title="Delete" data-id="${book.id}">🗑️</button>
                </div>
            </td>
        `;

        row.addEventListener('click', (e) => {
            if (e.target.closest('.quick-btn')) return;
            openBookDetailModal(book.id);
        });

        row.querySelector('.quick-edit').addEventListener('click', (e) => {
            e.stopPropagation();
            openEditModal(book.id);
        });

        row.querySelector('.quick-delete').addEventListener('click', (e) => {
            e.stopPropagation();
            openConfirmDeleteModal(book.id);
        });

        tbody.appendChild(row);
    });
}

// Render Analytics Tab
function renderAnalytics(data) {
    const total = data.length;
    if (total === 0) {
        document.getElementById('price-analytics').innerHTML = '<p>No data</p>';
        return;
    }

    // Price Stats
    const prices = data.map(d => parseFloat(d.Price));
    const priceSum = prices.reduce((a,b)=>a+b, 0);
    const priceAvg = priceSum / total;
    const priceMax = Math.max(...prices);
    const priceMin = Math.min(...prices);
    prices.sort((a,b)=>a-b);
    const priceMedian = prices[Math.floor(total/2)];

    document.getElementById('price-analytics').innerHTML = `
        <table class="analytics-table">
            <tr><td class="label-cell">Average Price</td><td class="value-cell">£${priceAvg.toFixed(2)}</td></tr>
            <tr><td class="label-cell">Median Price</td><td class="value-cell">£${priceMedian.toFixed(2)}</td></tr>
            <tr><td class="label-cell">Highest Price</td><td class="value-cell">£${priceMax.toFixed(2)}</td></tr>
            <tr><td class="label-cell">Lowest Price</td><td class="value-cell">£${priceMin.toFixed(2)}</td></tr>
            <tr><td class="label-cell">Total Value</td><td class="value-cell">£${priceSum.toFixed(2)}</td></tr>
        </table>
    `;

    // Rating Stats
    const ratings = data.map(d => d['Numeric Rating']);
    const ratingAvg = ratings.reduce((a,b)=>a+b, 0) / total;
    const ratingCounts = {1:0, 2:0, 3:0, 4:0, 5:0};
    ratings.forEach(r => ratingCounts[r]++);

    document.getElementById('rating-analytics').innerHTML = `
        <table class="analytics-table">
            <tr><td class="label-cell">Average Rating</td><td class="value-cell">${ratingAvg.toFixed(2)} ★</td></tr>
            <tr><td class="label-cell">5-Star Books</td><td class="value-cell">${ratingCounts[5]} (${Math.round(ratingCounts[5]/total*100)}%)</td></tr>
            <tr><td class="label-cell">4-Star Books</td><td class="value-cell">${ratingCounts[4]} (${Math.round(ratingCounts[4]/total*100)}%)</td></tr>
            <tr><td class="label-cell">3-Star Books</td><td class="value-cell">${ratingCounts[3]} (${Math.round(ratingCounts[3]/total*100)}%)</td></tr>
            <tr><td class="label-cell">2-Star Books</td><td class="value-cell">${ratingCounts[2]} (${Math.round(ratingCounts[2]/total*100)}%)</td></tr>
            <tr><td class="label-cell">1-Star Books</td><td class="value-cell">${ratingCounts[1]} (${Math.round(ratingCounts[1]/total*100)}%)</td></tr>
        </table>
    `;

    // Sentiment Breakdown
    const sentimentCounts = { 'Positive': 0, 'Neutral': 0, 'Negative': 0 };
    const polaritySum = data.reduce((sum, d) => sum + (parseFloat(d.Sentiment_Polarity) || 0), 0);
    const polarityAvg = polaritySum / total;
    data.forEach(d => {
        if (sentimentCounts[d.Sentiment_Class] !== undefined) sentimentCounts[d.Sentiment_Class]++;
    });

    document.getElementById('sentiment-analytics').innerHTML = `
        <table class="analytics-table">
            <tr><td class="label-cell">Average Polarity Score</td><td class="value-cell">${polarityAvg.toFixed(3)}</td></tr>
            <tr><td class="label-cell">Positive Sentiments</td><td class="value-cell" style="color:#10b981">${sentimentCounts['Positive']} (${Math.round(sentimentCounts['Positive']/total*100)}%)</td></tr>
            <tr><td class="label-cell">Neutral Sentiments</td><td class="value-cell" style="color:#94a3b8">${sentimentCounts['Neutral']} (${Math.round(sentimentCounts['Neutral']/total*100)}%)</td></tr>
            <tr><td class="label-cell">Negative Sentiments</td><td class="value-cell" style="color:#ef4444">${sentimentCounts['Negative']} (${Math.round(sentimentCounts['Negative']/total*100)}%)</td></tr>
        </table>
    `;

    // Emotion Distribution Analytics
    const emotionCounts = { 'Joy': 0, 'Surprise': 0, 'Disappointment': 0, 'Sadness': 0, 'Anger': 0, 'Neutral': 0 };
    data.forEach(d => {
        const emo = d.Emotion || 'Neutral';
        if (emotionCounts[emo] !== undefined) emotionCounts[emo]++;
    });

    document.getElementById('emotion-analytics').innerHTML = `
        <table class="analytics-table">
            <tr><td class="label-cell">Joy</td><td class="value-cell" style="color:#10b981">${emotionCounts['Joy']} (${Math.round(emotionCounts['Joy']/total*100)}%)</td></tr>
            <tr><td class="label-cell">Surprise</td><td class="value-cell" style="color:#a78bfa">${emotionCounts['Surprise']} (${Math.round(emotionCounts['Surprise']/total*100)}%)</td></tr>
            <tr><td class="label-cell">Disappointment</td><td class="value-cell" style="color:#fbbf24">${emotionCounts['Disappointment']} (${Math.round(emotionCounts['Disappointment']/total*100)}%)</td></tr>
            <tr><td class="label-cell">Sadness</td><td class="value-cell" style="color:#60a5fa">${emotionCounts['Sadness']} (${Math.round(emotionCounts['Sadness']/total*100)}%)</td></tr>
            <tr><td class="label-cell">Anger</td><td class="value-cell" style="color:#ef4444">${emotionCounts['Anger']} (${Math.round(emotionCounts['Anger']/total*100)}%)</td></tr>
            <tr><td class="label-cell">Neutral</td><td class="value-cell" style="color:#94a3b8">${emotionCounts['Neutral']} (${Math.round(emotionCounts['Neutral']/total*100)}%)</td></tr>
        </table>
    `;

    // Top Rated Books (Sorted by rating desc, then price asc)
    const topRated = [...data].sort((a,b) => {
        if(b['Numeric Rating'] !== a['Numeric Rating']) {
            return b['Numeric Rating'] - a['Numeric Rating'];
        }
        return parseFloat(a.Price) - parseFloat(b.Price);
    }).slice(0, 5);

    let topBooksHtml = '';
    topRated.forEach(b => {
        topBooksHtml += `
            <div class="analytics-list-item">
                <div class="item-left">
                    <span class="item-title">${b.Title}</span>
                    <span class="item-subtitle">${b.Author} | ${b.Genre || 'Unknown'}</span>
                </div>
                <div class="item-right" style="color:#fbbf24">${'★'.repeat(b['Numeric Rating'])}</div>
            </div>
        `;
    });
    document.getElementById('top-books').innerHTML = topBooksHtml || '<p>No data</p>';

    // Most Expensive Books
    const expensive = [...data].sort((a,b)=>parseFloat(b.Price)-parseFloat(a.Price)).slice(0, 5);
    let expBooksHtml = '';
    expensive.forEach(b => {
        expBooksHtml += `
            <div class="analytics-list-item">
                <div class="item-left">
                    <span class="item-title">${b.Title}</span>
                    <span class="item-subtitle">${b.Author}</span>
                </div>
                <div class="item-right">£${parseFloat(b.Price).toFixed(2)}</div>
            </div>
        `;
    });
    document.getElementById('expensive-books').innerHTML = expBooksHtml || '<p>No data</p>';

    // Genre Breakdown
    const genresMap = {};
    data.forEach(d => {
        const genre = d.Genre || 'Unknown';
        genresMap[genre] = (genresMap[genre] || 0) + 1;
    });
    const sortedGenres = Object.entries(genresMap).sort((a,b)=>b[1]-a[1]);
    let genreHtml = '<table class="analytics-table">';
    sortedGenres.slice(0, 5).forEach(([g, count]) => {
        genreHtml += `
            <tr>
                <td class="label-cell">${g}</td>
                <td class="value-cell">${count} books (${Math.round(count/total*100)}%)</td>
            </tr>
        `;
    });
    genreHtml += '</table>';
    document.getElementById('genre-analytics').innerHTML = genreHtml;
}

// Modal handling logic
function setupModals() {
    const bookModal = document.getElementById('book-modal');
    const editModal = document.getElementById('edit-modal');
    const addModal = document.getElementById('add-modal');
    const confirmModal = document.getElementById('confirm-modal');

    // Close buttons
    const closeBtns = [
        { btn: 'modal-close-btn', modal: bookModal },
        { btn: 'modal-cancel-btn', modal: bookModal },
        { btn: 'edit-modal-close', modal: editModal },
        { btn: 'edit-cancel-btn', modal: editModal },
        { btn: 'add-modal-close', modal: addModal },
        { btn: 'add-cancel-btn', modal: addModal },
        { btn: 'confirm-delete-no', modal: confirmModal }
    ];

    closeBtns.forEach(item => {
        const element = document.getElementById(item.btn);
        if (element) {
            element.addEventListener('click', () => {
                item.modal.classList.remove('active');
            });
        }
    });

    // Open add book modal
    document.getElementById('open-add-modal-btn').addEventListener('click', () => {
        document.getElementById('add-book-form').reset();
        addModal.classList.add('active');
    });

    // Close when clicking overlay
    [bookModal, editModal, addModal, confirmModal].forEach(modal => {
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.classList.remove('active');
            }
        });
    });
}

// Form Submission handling (Save & Add)
function setupForms() {
    // Edit Form Save
    const editForm = document.getElementById('edit-book-form');
    editForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const id = parseInt(document.getElementById('edit-id').value);
        const book = booksData.find(b => b.id === id);
        
        if (book) {
            book.Title = document.getElementById('edit-title').value;
            book.Author = document.getElementById('edit-author').value;
            book.Genre = document.getElementById('edit-genre').value;
            book.PublishYear = parseInt(document.getElementById('edit-year').value) || 2024;
            book.Pages = parseInt(document.getElementById('edit-pages').value) || 250;
            book.Publisher = document.getElementById('edit-publisher').value;
            book.Price = parseFloat(document.getElementById('edit-price').value);
            book['Numeric Rating'] = parseInt(document.getElementById('edit-rating').value);
            book.Review_Text = document.getElementById('edit-review').value;
            book.Description = document.getElementById('edit-description').value;

            // Recalculate sentiment locally
            const sent = estimateSentiment(book.Review_Text, book['Numeric Rating']);
            book.Sentiment_Polarity = sent.score;
            book.Sentiment_Class = sent.class;
            book.Emotion = sent.emotion;

            document.getElementById('edit-modal').classList.remove('active');
            renderDashboard();
            showToast(`💾 Saved changes to "${book.Title}"`);
        }
    });

    // Add Form Submit
    const addForm = document.getElementById('add-book-form');
    addForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const maxId = booksData.length > 0 ? Math.max(...booksData.map(b => b.id)) : 0;
        
        // Random covers palettes
        const coverColors = ['#6366f1', '#ec4899', '#f59e0b', '#dc2626', '#0891b2', '#9333ea', '#16a34a', '#be185d', '#2563eb', '#7c3aed', '#115e59'];
        const randomColor = coverColors[Math.floor(Math.random() * coverColors.length)];

        const rating = parseInt(document.getElementById('add-rating').value);
        const reviewText = document.getElementById('add-review').value;
        const sent = estimateSentiment(reviewText, rating);

        const newBook = {
            id: maxId + 1,
            Title: document.getElementById('add-title').value,
            Author: document.getElementById('add-author').value,
            Genre: document.getElementById('add-genre').value || 'Uncategorized',
            PublishYear: parseInt(document.getElementById('add-year').value) || 2024,
            PublishDate: new Date().toISOString().split('T')[0],
            Pages: parseInt(document.getElementById('add-pages').value) || 200,
            Publisher: document.getElementById('add-publisher').value || 'Self-Published',
            ISBN: '978-' + Math.floor(Math.random() * 10000000000),
            Description: document.getElementById('add-description').value || 'No synopsis provided.',
            CoverColor: randomColor,
            'Numeric Rating': rating,
            Price: parseFloat(document.getElementById('add-price').value),
            Review_Text: reviewText || 'No review comments.',
            Sentiment_Polarity: sent.score,
            Sentiment_Class: sent.class,
            Emotion: sent.emotion
        };

        booksData.push(newBook);
        document.getElementById('add-modal').classList.remove('active');
        renderDashboard();
        showToast(`➕ Added "${newBook.Title}" to catalog!`);
    });

    // Action handlers inside book detail modal
    document.getElementById('modal-edit-btn').addEventListener('click', () => {
        if (selectedBookId) {
            document.getElementById('book-modal').classList.remove('active');
            openEditModal(selectedBookId);
        }
    });

    document.getElementById('modal-delete-btn').addEventListener('click', () => {
        if (selectedBookId) {
            document.getElementById('book-modal').classList.remove('active');
            openConfirmDeleteModal(selectedBookId);
        }
    });

    // Confirm Delete Action
    document.getElementById('confirm-delete-yes').addEventListener('click', () => {
        if (selectedBookId) {
            const index = booksData.findIndex(b => b.id === selectedBookId);
            if (index !== -1) {
                const title = booksData[index].Title;
                booksData.splice(index, 1);
                document.getElementById('confirm-modal').classList.remove('active');
                renderDashboard();
                showToast(`🗑️ Deleted "${title}"`);
            }
        }
    });
}

// Open Book Details Modal
function openBookDetailModal(id) {
    const book = booksData.find(b => b.id === id);
    if (!book) return;

    selectedBookId = id;
    
    // Set up cover rendering
    const coverElement = document.getElementById('modal-cover');
    const color = book.CoverColor || '#4b5563';
    coverElement.style.background = `linear-gradient(135deg, ${color}, #1e293b)`;
    document.getElementById('modal-cover-title').textContent = book.Title;
    document.getElementById('modal-cover-author').textContent = book.Author;

    // Set Text Fields
    document.getElementById('modal-genre').textContent = book.Genre || 'Unknown';
    document.getElementById('modal-title').textContent = book.Title;
    document.getElementById('modal-author').textContent = `by ${book.Author}`;
    document.getElementById('modal-year').textContent = book.PublishYear || 'N/A';
    document.getElementById('modal-pages').textContent = book.Pages ? `${book.Pages} pages` : 'N/A';
    document.getElementById('modal-publisher').textContent = book.Publisher || 'N/A';
    document.getElementById('modal-isbn').textContent = book.ISBN || 'N/A';
    document.getElementById('modal-price').textContent = `£${parseFloat(book.Price).toFixed(2)}`;
    
    // Stars Rating
    const starsHtml = '★'.repeat(book['Numeric Rating']) + '☆'.repeat(5 - book['Numeric Rating']);
    document.getElementById('modal-rating-stars').innerHTML = `<span style="color:#fbbf24">${starsHtml}</span>`;

    document.getElementById('modal-description').textContent = book.Description || 'No synopsis available.';
    document.getElementById('modal-review').textContent = book.Review_Text || 'No review comments.';

    // Sentiment badge
    const badge = document.getElementById('modal-sentiment-badge');
    badge.className = `badge ${book.Sentiment_Class.toLowerCase()}`;
    badge.textContent = `${book.Sentiment_Class} (${book.Sentiment_Polarity > 0 ? '+' : ''}${book.Sentiment_Polarity})`;

    const emotionBadge = document.getElementById('modal-emotion-badge');
    if (emotionBadge) {
        const emo = book.Emotion || 'Neutral';
        emotionBadge.className = `badge ${emo.toLowerCase()}`;
        emotionBadge.textContent = emo;
        emotionBadge.style.display = 'inline-block';
    }

    document.getElementById('book-modal').classList.add('active');
}

// Open Edit Book Modal
function openEditModal(id) {
    const book = booksData.find(b => b.id === id);
    if (!book) return;

    selectedBookId = id;
    
    document.getElementById('edit-id').value = book.id;
    document.getElementById('edit-title').value = book.Title;
    document.getElementById('edit-author').value = book.Author;
    document.getElementById('edit-genre').value = book.Genre || '';
    document.getElementById('edit-year').value = book.PublishYear || '';
    document.getElementById('edit-pages').value = book.Pages || '';
    document.getElementById('edit-publisher').value = book.Publisher || '';
    document.getElementById('edit-price').value = parseFloat(book.Price).toFixed(2);
    document.getElementById('edit-rating').value = book['Numeric Rating'];
    document.getElementById('edit-review').value = book.Review_Text || '';
    document.getElementById('edit-description').value = book.Description || '';

    document.getElementById('edit-modal').classList.add('active');
}

// Open Delete Confirmation Modal
function openConfirmDeleteModal(id) {
    const book = booksData.find(b => b.id === id);
    if (!book) return;

    selectedBookId = id;
    document.getElementById('confirm-delete-name').textContent = `"${book.Title}"`;
    document.getElementById('confirm-modal').classList.add('active');
}

// Helper to show Toast messages
function showToast(message) {
    const toast = document.getElementById('toast');
    toast.innerHTML = `<span>🔔</span> ${message}`;
    toast.classList.add('show');
    
    setTimeout(() => {
        toast.classList.remove('show');
    }, 3500);
}
