// 現在の日時を取得
const updateDateTIme = () => {
    const now = new Date();
    const datetimeString = now.toLocaleString('ja-JP', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: "2-digit",
        hour12: false, // 24時間制
    });
    document.getElementById('datetime').textContent = '現在日時：' + datetimeString;
};

// 初回表示
updateDateTIme();

// 1秒毎に更新
setInterval(updateDateTIme, 1000);