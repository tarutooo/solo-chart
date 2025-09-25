document.addEventListener('DOMContentLoaded', function() {
    // Load tasks from localStorage (fallback to cookie) or use default tasks
    let tasks = loadTasksFromStorage() || loadTasksFromCookie() || [
        { id: 1, name: 'Task 1', start: '2025-06-04', end: '2025-06-08', color: '#7DB9DE', memo: 'これは最初のタスクです。' },
        { id: 2, name: 'Task 2', start: '2025-06-06', end: '2025-06-12', color: '#F7A072', memo: '' },
        { id: 3, name: 'Task 3', start: '2025-06-09', end: '2025-06-15', color: '#A3DE83', memo: '詳細な説明' },
    ];
    let nextId = tasks.length > 0 ? Math.max(...tasks.map(t => t.id)) + 1 : 1;

    // 今日のローカル日付(YYYY-MM-DD)を取得
    const todayISO = toLocalISODate(new Date());
    // タスク塗りの固定色
    const FIXED_TASK_COLOR = '#edff84ff';
    // 初期データを過去日不可のルールに合わせて丸める
    sanitizeTasksNotBeforeToday();

    const ganttChartContainer = document.getElementById('gantt-chart-container');
    const taskTableBody = document.getElementById('task-table-body');
    const addRowBtn = document.getElementById('add-row-btn');
    const resetBtn = document.getElementById('reset-btn');

    // Storage functions (localStorage primary, cookie fallback for legacy)
    function saveTasksToStorage() {
        try {
            localStorage.setItem('ganttTasksV2', JSON.stringify(tasks));
        } catch (e) {
            console.warn('localStorage save failed, falling back to cookie.', e);
            saveTasksToCookie();
        }
        // 互換のためクッキーにも保存（容量上限に注意）
        saveTasksToCookie();
    }

    function loadTasksFromStorage() {
        try {
            const s = localStorage.getItem('ganttTasksV2');
            if (s) return JSON.parse(s);
        } catch (e) {
            console.warn('localStorage load failed.', e);
        }
        return null;
    }

    function saveTasksToCookie() {
        try {
            const cookieData = JSON.stringify(tasks);
            const expirationDate = new Date();
            expirationDate.setTime(expirationDate.getTime() + (30 * 24 * 60 * 60 * 1000)); // 30 days
            document.cookie = `ganttTasks=${encodeURIComponent(cookieData)}; expires=${expirationDate.toUTCString()}; path=/`;
        } catch (_) { /* no-op */ }
    }

    function loadTasksFromCookie() {
        const cookies = document.cookie.split(';');
        for (let cookie of cookies) {
            const [name, value] = cookie.trim().split('=');
            if (name === 'ganttTasks') {
                try {
                    return JSON.parse(decodeURIComponent(value));
                } catch (error) {
                    console.error('Error parsing tasks from cookie:', error);
                    return null;
                }
            }
        }
        return null;
    }

    // Scroll position persistence
    function saveScrollPosition() {
        try {
            const container = document.getElementById('gantt-chart-container');
            if (container) localStorage.setItem('ganttScrollLeft', String(container.scrollLeft));
        } catch (_) { /* no-op */ }
    }

    function restoreScrollPosition() {
        try {
            const container = document.getElementById('gantt-chart-container');
            if (!container) return;
            const saved = localStorage.getItem('ganttScrollLeft');
            if (saved !== null) {
                container.scrollLeft = parseInt(saved, 10) || 0;
            }
        } catch (_) { /* no-op */ }
    }

    function renderSpreadsheet() {
        taskTableBody.innerHTML = ''; // Clear existing rows
        tasks.forEach(task => {
            const row = taskTableBody.insertRow();
            row.setAttribute('data-task-id', task.id);

            // Task Name Cell
            const nameCell = row.insertCell();
            const nameInput = document.createElement('textarea');
            nameInput.value = task.name;
            nameInput.rows = 2;
            nameInput.placeholder = 'タスク名';
            nameInput.addEventListener('change', (e) => updateTask(task.id, 'name', e.target.value));
            nameInput.addEventListener('input', debounce((e) => updateTask(task.id, 'name', e.target.value), 200));
            nameCell.appendChild(nameInput);

            // Start Date Cell
            const startCell = row.insertCell();
            const startInput = document.createElement('input');
            startInput.type = 'date';
            startInput.value = task.start;
            startInput.min = todayISO; // 過去日は選択不可
            startInput.addEventListener('change', (e) => updateTask(task.id, 'start', e.target.value));
            startInput.addEventListener('input', debounce((e) => updateTask(task.id, 'start', e.target.value), 200));
            startCell.appendChild(startInput);

            // End Date Cell
            const endCell = row.insertCell();
            const endInput = document.createElement('input');
            endInput.type = 'date';
            endInput.value = task.end;
            // 終了日の下限は今日 or 開始日のどちらか遅い方
            endInput.min = (task.start && task.start > todayISO) ? task.start : todayISO;
            endInput.addEventListener('change', (e) => updateTask(task.id, 'end', e.target.value));
            endInput.addEventListener('input', debounce((e) => updateTask(task.id, 'end', e.target.value), 200));
            endCell.appendChild(endInput);

            // Duration (Days) Cell - 読み取り専用
            const daysCell = row.insertCell();
            daysCell.className = 'days-cell';
            daysCell.textContent = formatDaysCount(task.start, task.end);

            // Memo Cell
            const memoCell = row.insertCell();
            const memoInput = document.createElement('textarea');
            memoInput.value = task.memo || '';
            memoInput.rows = 3;
            memoInput.placeholder = 'メモを書いてください';
            memoInput.addEventListener('change', (e) => updateTask(task.id, 'memo', e.target.value));
            memoInput.addEventListener('input', debounce((e) => updateTask(task.id, 'memo', e.target.value), 200));
            memoCell.appendChild(memoInput);

            // Action Cell (Delete Button)
            const actionCell = row.insertCell();
            const deleteBtn = document.createElement('button');
            deleteBtn.textContent = '削除';
            deleteBtn.className = 'delete-btn';
            deleteBtn.addEventListener('click', () => deleteTask(task.id));
            actionCell.appendChild(deleteBtn);
        });
    }

    function updateTask(id, field, value) {
        const task = tasks.find(t => t.id === id);
        if (task) {
            if (field === 'start' || field === 'end') {
                // 日付フィールドのみ制約適用
                let v = value;
                if (v < todayISO) v = todayISO;
                task[field] = v;
                // 整合性
                if (field === 'start') {
                    if (!task.end || task.end < task.start) task.end = task.start;
                } else if (field === 'end') {
                    if (!task.start || task.start > task.end) task.start = task.end;
                }
                // End min 更新
                updateEndMin(task.id);
                // 日数セル更新
                updateDaysCell(task.id);
            } else {
                // テキスト系はそのまま
                task[field] = value;
            }
            renderGanttChart();
            saveTasksToStorage();
        }
    }

    function updateEndMin(taskId) {
        const row = taskTableBody.querySelector(`tr[data-task-id="${taskId}"]`);
        if (!row) return;
        const inputs = row.querySelectorAll('input[type="date"]');
        const t = tasks.find(x => x.id === taskId);
        if (inputs[1] && t) {
            inputs[1].min = (t.start && t.start > todayISO) ? t.start : todayISO;
        }
    }

    function daysBetweenInclusive(sISO, eISO) {
        if (!sISO || !eISO) return 0;
        const s = new Date(sISO);
        const e = new Date(eISO);
        const one = new Date(s.getFullYear(), s.getMonth(), s.getDate());
        const two = new Date(e.getFullYear(), e.getMonth(), e.getDate());
        const diff = Math.round((two - one) / 86400000);
        return diff + 1; // 両端含む
    }

    function formatDaysCount(sISO, eISO) {
        const d = daysBetweenInclusive(sISO, eISO);
        return Number.isFinite(d) && d > 0 ? `${d}日` : '-';
    }

    function updateDaysCell(taskId) {
        const row = taskTableBody.querySelector(`tr[data-task-id="${taskId}"]`);
        if (!row) return;
        const t = tasks.find(x => x.id === taskId);
        const idx = 3; // タスク名,開始,終了,日数,メモ,アクション
        if (row.cells[idx]) row.cells[idx].textContent = formatDaysCount(t?.start, t?.end);
    }

    function addNewTask() {
        const today = todayISO; // ローカル日付で揃える
        const newTask = {
            id: nextId++,
            name: '新しいタスク',
            start: today,
            end: today,
            memo: '',
            color: getRandomColor()
        };
        tasks.push(newTask);
        renderSpreadsheet();
        renderGanttChart();
    saveTasksToStorage();
    }

    function deleteTask(id) {
        tasks = tasks.filter(t => t.id !== id);
        renderSpreadsheet();
        renderGanttChart();
    saveTasksToStorage();
    }

    function renderGanttChart() {
        ganttChartContainer.innerHTML = '';
        if (tasks.length === 0) return;

        const table = document.createElement('table');
        table.className = 'gantt-table';

        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');

        // --- ヘッダーの生成（3段: 月 / 週 / 日） ---
        const monthRow = document.createElement('tr');
        monthRow.className = 'month-row';
        const weekRow = document.createElement('tr');
        weekRow.className = 'week-row';
        const dayRow = document.createElement('tr');
        dayRow.className = 'day-row';

        const taskNameHeader = document.createElement('th');
        taskNameHeader.className = 'gantt-task-name';
        taskNameHeader.innerText = 'タスク';
        taskNameHeader.rowSpan = 3;
        monthRow.appendChild(taskNameHeader);

        const dates = getChartDates();

        // 月ヘッダー
        let i = 0;
        while (i < dates.length) {
            const y = dates[i].getFullYear();
            const m = dates[i].getMonth();
            let span = 1;
            while (i + span < dates.length &&
                   dates[i + span].getFullYear() === y &&
                   dates[i + span].getMonth() === m) {
                span++;
            }
            const th = document.createElement('th');
            th.className = 'month';
            th.colSpan = span;
            th.innerText = `${y}/${m + 1}`;
            monthRow.appendChild(th);
            i += span;
        }

        // 週ヘッダー（各月で 週1, 週2 ... とリセット。週の起点=月曜）
        i = 0;
        while (i < dates.length) {
            const base = dates[i];
            const month = base.getMonth();
            const weekInMonth = weekOfMonthMonday(base);
            let span = 1;
            while (i + span < dates.length) {
                const dnext = dates[i + span];
                if (dnext.getMonth() !== month) break;
                if (weekOfMonthMonday(dnext) !== weekInMonth) break;
                span++;
            }
            const th = document.createElement('th');
            th.className = 'week';
            th.colSpan = span;
            th.innerText = `週${weekInMonth}`;
            weekRow.appendChild(th);
            i += span;
        }

        // 日ヘッダー（全日ラベル表示 + 強調クラス）
        dates.forEach((date, idx) => {
            const th = document.createElement('th');
            th.classList.add('day'); // 固定幅指定に対応
            const dow = date.getDay();
            const isWknd = dow === 0 || dow === 6;
            const isWkStart = dow === 1;
            const isMonthStart = date.getDate() === 1;
            if (isWknd) th.classList.add('weekend');
            if (isWkStart) th.classList.add('week-start');
            if (isMonthStart) th.classList.add('month-start');
            if (isSameDay(date, new Date())) th.classList.add('today');
            const dowShort = ['日','月','火','水','木','金','土'][dow];
            th.innerHTML = `<span class="dow">${dowShort}</span><span class="dom">${date.getDate()}</span>`;
            dayRow.appendChild(th);
        });

        thead.appendChild(monthRow);
        thead.appendChild(weekRow);
        thead.appendChild(dayRow);

        // --- ボディの生成 ---
        tasks.forEach(task => {
            const taskRow = document.createElement('tr');
            const taskNameCell = document.createElement('td');
            taskNameCell.className = 'gantt-task-name';
            taskNameCell.innerText = task.name;
            taskRow.appendChild(taskNameCell);

            dates.forEach(date => {
                const cell = document.createElement('td');
                const dow = date.getDay();
                if (dow === 0 || dow === 6) cell.classList.add('weekend');
                if (dow === 1) cell.classList.add('week-start');
                if (date.getDate() === 1) cell.classList.add('month-start');
                if (isSameDay(date, new Date())) cell.classList.add('today');
                const taskStart = new Date(task.start);
                const taskEnd = new Date(task.end);
                
                // 日付の比較は年月日のみで行う
                const d1 = new Date(date.getFullYear(), date.getMonth(), date.getDate());
                const s = new Date(taskStart.getFullYear(), taskStart.getMonth(), taskStart.getDate());
                const e = new Date(taskEnd.getFullYear(), taskEnd.getMonth(), taskEnd.getDate());

                if (d1 >= s && d1 <= e) {
                    // 固定色で塗る。background を使って weekend/today の背景イメージを上書き
                    cell.style.background = FIXED_TASK_COLOR;
                    cell.classList.add('filled');
                }
                taskRow.appendChild(cell);
            });
            tbody.appendChild(taskRow);
        });

        table.appendChild(thead);
        table.appendChild(tbody);
    ganttChartContainer.appendChild(table);
    // スクロール位置復元（再描画時も維持）
    restoreScrollPosition();
    }

    function getChartDates() {
        if (tasks.length === 0) return [];

        const startDates = tasks.map(t => new Date(t.start));
        const endDates = tasks.map(t => new Date(t.end));

        let minDate = new Date(Math.min.apply(null, startDates));
        let maxDate = new Date(Math.max.apply(null, endDates));
        
        // 期間が短すぎる場合、最低でも1ヶ月表示する
        const oneMonthLater = new Date(minDate);
        oneMonthLater.setMonth(oneMonthLater.getMonth() + 1);
        if (maxDate < oneMonthLater) {
            maxDate = oneMonthLater;
        }


        let dates = [];
        let currentDate = new Date(minDate);
        currentDate.setDate(currentDate.getDate() - 2); // 開始日の2日前から表示
        maxDate.setDate(maxDate.getDate() + 2); // 終了日の2日後まで表示

        while (currentDate <= maxDate) {
            dates.push(new Date(currentDate));
            currentDate.setDate(currentDate.getDate() + 1);
        }
        return dates;
    }


    function isWeekend(date) {
        const day = date.getDay();
        return day === 0 || day === 6; // Sun or Sat
    }

    function isSameDay(a, b) {
        return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
    }

    // ISO week number
    function isoWeekNumber(d) {
        const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
        // Thursday in current week decides the year.
        const dayNum = date.getUTCDay() || 7;
        date.setUTCDate(date.getUTCDate() + 4 - dayNum);
        const yearStart = new Date(Date.UTC(date.getUTCFullYear(),0,1));
        const weekNo = Math.ceil((((date - yearStart) / 86400000) + 1)/7);
        return { week: weekNo, weekYear: date.getUTCFullYear() };
    }

    function getRandomColor() {
        const letters = '0123456789ABCDEF';
        let color = '#';
        for (let i = 0; i < 6; i++) {
            color += letters[Math.floor(Math.random() * 16)];
        }
        return color;
    }

    function darkenColor(color, percent) {
        let num = parseInt(color.replace("#",""), 16),
        amt = Math.round(2.55 * percent),
        R = (num >> 16) - amt,
        G = (num >> 8 & 0x00FF) - amt,
        B = (num & 0x0000FF) - amt;
        R = Math.max(0, Math.min(255, R));
        G = Math.max(0, Math.min(255, G));
        B = Math.max(0, Math.min(255, B));
        return "#" + (0x1000000 + R*0x10000 + G*0x100 + B).toString(16).slice(1);
    }

    // 月曜は週の開始として、月内の週番号(1..)を返す
    function weekOfMonthMonday(d) {
        const firstOfMonth = new Date(d.getFullYear(), d.getMonth(), 1);
        const firstDow = firstOfMonth.getDay(); // 0(Sun)..6(Sat)
        const mondayOffset = (firstDow + 6) % 7; // Mon=0, Tue=1, ... Sun=6
        return Math.floor((d.getDate() + mondayOffset - 1) / 7) + 1;
    }

    // ローカルタイムゾーンで YYYY-MM-DD を返す
    function toLocalISODate(d) {
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${day}`;
    }

    // 初期データの過去日を今日へ丸め、整合性を取る
    function sanitizeTasksNotBeforeToday() {
        tasks.forEach(t => {
            if (t.start && t.start < todayISO) t.start = todayISO;
            if (t.end && t.end < todayISO) t.end = todayISO;
            if (t.start && t.end && t.end < t.start) t.end = t.start;
        });
    }

    // Event Listeners
    addRowBtn.addEventListener('click', addNewTask);
    resetBtn.addEventListener('click', () => {
        if (confirm('すべてのタスクをリセットしますか？')) {
            document.cookie = "ganttTasks=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;";
            tasks = [];
            nextId = 1;
        try { localStorage.removeItem('ganttTasksV2'); } catch (_) {}
        try { localStorage.removeItem('ganttScrollLeft'); } catch (_) {}
            renderSpreadsheet();
            renderGanttChart();
        }
    });

    // Initial render
    renderSpreadsheet();
    renderGanttChart();

    // スクロール保存イベント
    const container = document.getElementById('gantt-chart-container');
    if (container) container.addEventListener('scroll', saveScrollPosition, { passive: true });

    // リロード/タブ非表示時に未保存の編集をフラッシュ
    window.addEventListener('beforeunload', () => { flushInputsToTasks(); saveTasksToStorage(); saveScrollPosition(); });
    document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'hidden') { flushInputsToTasks(); saveTasksToStorage(); saveScrollPosition(); }
    });

    // 入力欄から現在の値を読み取り、tasksに反映して保存
    function flushInputsToTasks() {
        const rows = taskTableBody.querySelectorAll('tr');
        rows.forEach(r => {
            const id = Number(r.getAttribute('data-task-id'));
            const t = tasks.find(x => x.id === id);
            if (!t) return;
            const nameEl = r.querySelector('td:first-child textarea') || r.querySelector('td:first-child input[type="text"]');
            const dateEls = r.querySelectorAll('input[type="date"]');
            // メモ欄は5列目に固定されているため、列指定で確実に取得する
            const memoEl = r.querySelector('td:nth-child(5) textarea');
            if (nameEl) t.name = nameEl.value;
            if (dateEls[0]) t.start = (dateEls[0].value && dateEls[0].value >= todayISO) ? dateEls[0].value : todayISO;
            if (dateEls[1]) t.end = (dateEls[1].value && dateEls[1].value >= todayISO) ? dateEls[1].value : (t.start || todayISO);
            if (t.start && t.end && t.end < t.start) t.end = t.start;
            if (memoEl) t.memo = memoEl.value || '';
        });
    }

    // 短い間隔での保存連打を避ける
    function debounce(fn, wait) {
        let to;
        return function(...args) {
            clearTimeout(to);
            to = setTimeout(() => fn.apply(this, args), wait);
        };
    }

    // タブ機能の初期化
    initTabFunctionality();

    // インクリメンタルモデル機能の初期化
    initIncrementalModel();

    function initTabFunctionality() {
        const tabButtons = document.querySelectorAll('.tab-btn');
        const tabContents = document.querySelectorAll('.tab-content');

        tabButtons.forEach(button => {
            button.addEventListener('click', () => {
                const targetTab = button.getAttribute('data-tab');
                // タブ切り替え前に現在の入力をフラッシュして保存（特にメモ欄の喪失対策）
                try { flushInputsToTasks(); saveTasksToStorage(); } catch (_) {}
                
                // すべてのタブボタンからactiveクラスを削除
                tabButtons.forEach(btn => btn.classList.remove('active'));
                // すべてのタブコンテンツからactiveクラスを削除
                tabContents.forEach(content => content.classList.remove('active'));
                
                // クリックされたタブボタンにactiveクラスを追加
                button.classList.add('active');
                // 対応するタブコンテンツにactiveクラスを追加
                document.getElementById(`${targetTab}-tab`).classList.add('active');
                
                // ガントチャートタブがアクティブになった時にチャートを再描画
                if (targetTab === 'gantt') {
                    setTimeout(() => {
                        renderGanttChart();
                    }, 100);
                }
                
                // インクリメンタルモデルタブがアクティブになった時にタイムラインを再描画
                if (targetTab === 'incremental') {
                    setTimeout(() => {
                        renderIncrementTimeline();
                    }, 100);
                }
            });
        });
    }

    function initIncrementalModel() {
        // インクリメントデータの管理
        let increments = loadIncrementsFromStorage() || [];
        let nextIncrementId = increments.length > 0 ? Math.max(...increments.map(i => i.id)) + 1 : 1;

        const addIncrementBtn = document.getElementById('add-increment-btn');
        const incrementsContainer = document.getElementById('increments-container');
        const timelineContainer = document.getElementById('increment-timeline-container');
        
        // 順番変更用のボタン
        const moveUpBtn = document.getElementById('move-up-btn');
        const moveDownBtn = document.getElementById('move-down-btn');
        const clearSelectionBtn = document.getElementById('clear-selection-btn');
        
        let selectedIncrement = null;

        // インクリメント追加ボタンのイベントリスナー
        addIncrementBtn.addEventListener('click', () => {
            const name = document.getElementById('increment-name').value.trim();
            const start = document.getElementById('increment-start').value;
            const end = document.getElementById('increment-end').value;
            const goals = document.getElementById('increment-goals').value.trim();

            if (!name || !start || !end) {
                alert('名前、開始日、終了日は必須です。');
                return;
            }

            if (new Date(end) < new Date(start)) {
                alert('終了日は開始日以降に設定してください。');
                return;
            }

            const newIncrement = {
                id: nextIncrementId++,
                name: name,
                start: start,
                end: end,
                goals: goals,
                createdAt: new Date().toISOString()
            };

            increments.push(newIncrement);
            saveIncrementsToStorage();
            renderIncrements();
            renderIncrementTimeline();

            // フォームをクリア
            document.getElementById('increment-name').value = '';
            document.getElementById('increment-start').value = '';
            document.getElementById('increment-end').value = '';
            document.getElementById('increment-goals').value = '';
        });

        // 順番変更ボタンのイベントリスナー
        moveUpBtn.addEventListener('click', () => {
            if (selectedIncrement) {
                moveIncrementUp(selectedIncrement.id);
            }
        });

        moveDownBtn.addEventListener('click', () => {
            if (selectedIncrement) {
                moveIncrementDown(selectedIncrement.id);
            }
        });

        clearSelectionBtn.addEventListener('click', () => {
            clearSelection();
        });

        // インクリメント一覧の描画
        function renderIncrements() {
            incrementsContainer.innerHTML = '';

            if (increments.length === 0) {
                incrementsContainer.innerHTML = '<p style="color: #666; text-align: center; padding: 20px;">まだインクリメントが追加されていません。</p>';
                return;
            }

            // カスタム順序が設定されていない場合は開始日順でソート
            if (!increments.some(i => i.customOrder !== undefined)) {
                increments.sort((a, b) => new Date(a.start) - new Date(b.start));
                increments.forEach((increment, index) => {
                    increment.customOrder = index;
                });
            } else {
                // カスタム順序でソート
                increments.sort((a, b) => (a.customOrder || 0) - (b.customOrder || 0));
            }

            increments.forEach((increment, index) => {
                const card = document.createElement('div');
                card.className = 'increment-card';
                card.setAttribute('data-increment-id', increment.id);
                card.innerHTML = `
                    <div class="increment-number">${index + 1}</div>
                    <div class="increment-content">
                        <div class="increment-header">
                            <div class="increment-title">${escapeHtml(increment.name)}</div>
                            <div class="increment-dates">${formatDate(increment.start)} - ${formatDate(increment.end)}</div>
                        </div>
                        <div class="increment-goals">${escapeHtml(increment.goals || '目標・成果物が設定されていません')}</div>
                        <div class="increment-actions">
                            <button class="edit-increment-btn" onclick="editIncrement(${increment.id})">編集</button>
                            <button class="delete-increment-btn" onclick="deleteIncrement(${increment.id})">削除</button>
                        </div>
                    </div>
                `;
                
                // クリック選択イベントの設定
                card.addEventListener('click', (e) => {
                    // 編集・削除ボタンのクリック時は選択しない
                    if (e.target.classList.contains('edit-increment-btn') || 
                        e.target.classList.contains('delete-increment-btn')) {
                        return;
                    }
                    selectIncrement(increment, card);
                });
                
                incrementsContainer.appendChild(card);
            });
        }

        // インクリメント選択機能
        function selectIncrement(increment, cardElement) {
            // 既に選択されている場合は選択解除
            if (selectedIncrement && selectedIncrement.id === increment.id) {
                clearSelection();
                return;
            }

            // 前の選択を解除
            document.querySelectorAll('.increment-card').forEach(card => {
                card.classList.remove('selected');
            });

            // 新しい選択を設定
            selectedIncrement = increment;
            cardElement.classList.add('selected');
            updateButtonStates();
        }

        function clearSelection() {
            selectedIncrement = null;
            document.querySelectorAll('.increment-card').forEach(card => {
                card.classList.remove('selected');
            });
            updateButtonStates();
        }

        function updateButtonStates() {
            if (!selectedIncrement) {
                moveUpBtn.disabled = true;
                moveDownBtn.disabled = true;
                clearSelectionBtn.disabled = true;
                return;
            }

            const currentIndex = increments.findIndex(i => i.id === selectedIncrement.id);
            moveUpBtn.disabled = currentIndex <= 0;
            moveDownBtn.disabled = currentIndex >= increments.length - 1;
            clearSelectionBtn.disabled = false;
        }

        // インクリメントを上に移動
        function moveIncrementUp(incrementId) {
            const currentIndex = increments.findIndex(i => i.id === incrementId);
            if (currentIndex <= 0) return;

            // 配列内で要素を交換
            [increments[currentIndex], increments[currentIndex - 1]] = 
            [increments[currentIndex - 1], increments[currentIndex]];

            // カスタム順序を更新
            increments.forEach((increment, index) => {
                increment.customOrder = index;
            });

            saveIncrementsToStorage();
            renderIncrements();
            renderIncrementTimeline();
            
            // 選択を維持
            setTimeout(() => {
                const newCard = document.querySelector(`[data-increment-id="${incrementId}"]`);
                if (newCard) {
                    selectIncrement(selectedIncrement, newCard);
                }
            }, 50);
        }

        // インクリメントを下に移動
        function moveIncrementDown(incrementId) {
            const currentIndex = increments.findIndex(i => i.id === incrementId);
            if (currentIndex >= increments.length - 1) return;

            // 配列内で要素を交換
            [increments[currentIndex], increments[currentIndex + 1]] = 
            [increments[currentIndex + 1], increments[currentIndex]];

            // カスタム順序を更新
            increments.forEach((increment, index) => {
                increment.customOrder = index;
            });

            saveIncrementsToStorage();
            renderIncrements();
            renderIncrementTimeline();
            
            // 選択を維持
            setTimeout(() => {
                const newCard = document.querySelector(`[data-increment-id="${incrementId}"]`);
                if (newCard) {
                    selectIncrement(selectedIncrement, newCard);
                }
            }, 50);
        }

        // インクリメントタイムラインの描画
        function renderIncrementTimeline() {
            timelineContainer.innerHTML = '';

            if (increments.length === 0) {
                timelineContainer.innerHTML = '<p style="color: #666; text-align: center; padding: 20px;">インクリメントが追加されるとタイムラインが表示されます。</p>';
                return;
            }

            // カスタム順序またはデフォルト順序でソート
            const sortedIncrements = [...increments].sort((a, b) => {
                if (a.customOrder !== undefined && b.customOrder !== undefined) {
                    return a.customOrder - b.customOrder;
                }
                return new Date(a.start) - new Date(b.start);
            });

            sortedIncrements.forEach((increment, index) => {
                const startDate = new Date(increment.start);
                const endDate = new Date(increment.end);
                const duration = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;

                const timelineItem = document.createElement('div');
                timelineItem.className = 'timeline-item';
                timelineItem.style.animationDelay = `${index * 0.1}s`;
                timelineItem.innerHTML = `
                    <div class="timeline-date">
                        <strong>Phase ${index + 1}</strong><br>
                        ${formatDate(increment.start)}
                    </div>
                    <div class="timeline-content">
                        <div class="timeline-title">${escapeHtml(increment.name)}</div>
                        <div class="timeline-duration">${duration}日間 (${formatDate(increment.start)} - ${formatDate(increment.end)})</div>
                    </div>
                `;
                timelineContainer.appendChild(timelineItem);
            });
        }

        // インクリメント編集機能
        window.editIncrement = function(id) {
            const increment = increments.find(i => i.id === id);
            if (!increment) return;

            const name = prompt('インクリメント名:', increment.name);
            if (name === null) return;

            const start = prompt('開始日 (YYYY-MM-DD):', increment.start);
            if (start === null) return;

            const end = prompt('終了日 (YYYY-MM-DD):', increment.end);
            if (end === null) return;

            const goals = prompt('目標・成果物:', increment.goals);
            if (goals === null) return;

            if (!name.trim() || !start || !end) {
                alert('名前、開始日、終了日は必須です。');
                return;
            }

            if (new Date(end) < new Date(start)) {
                alert('終了日は開始日以降に設定してください。');
                return;
            }

            increment.name = name.trim();
            increment.start = start;
            increment.end = end;
            increment.goals = goals.trim();

            saveIncrementsToStorage();
            clearSelection();
            renderIncrements();
            renderIncrementTimeline();
        };

        // インクリメント削除機能
        window.deleteIncrement = function(id) {
            if (!confirm('このインクリメントを削除しますか？')) return;

            increments = increments.filter(i => i.id !== id);
            
            // 削除されたアイテムが選択されていた場合は選択解除
            if (selectedIncrement && selectedIncrement.id === id) {
                clearSelection();
            }
            
            saveIncrementsToStorage();
            renderIncrements();
            renderIncrementTimeline();
        };

        // インクリメントデータの保存・読み込み
        function saveIncrementsToStorage() {
            try {
                localStorage.setItem('incrementalModelData', JSON.stringify(increments));
            } catch (e) {
                console.warn('localStorage save failed for increments.', e);
            }
        }

        function loadIncrementsFromStorage() {
            try {
                const data = localStorage.getItem('incrementalModelData');
                return data ? JSON.parse(data) : null;
            } catch (e) {
                console.warn('localStorage load failed for increments.', e);
                return null;
            }
        }

        // 日付フォーマット関数
        function formatDate(dateString) {
            const date = new Date(dateString);
            return date.toLocaleDateString('ja-JP');
        }

        // HTML エスケープ関数
        function escapeHtml(text) {
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }

        // 初期描画
        renderIncrements();
        renderIncrementTimeline();
    }
});
