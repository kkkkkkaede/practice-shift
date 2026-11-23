"use strict";

/*
 * ICSファイルに出力する勤務情報
 *
 * 例:
 * {
 *   date: 2026-11-01,
 *   start: 2026-11-01 09:00,
 *   end: 2026-11-01 17:00
 * }
 */
let shiftEvents = [];


/*
 * シフト確認ボタン
 */
document
    .getElementById("checkButton")
    .addEventListener("click", async function () {
        const fileInput =
            document.getElementById("excelFile");

        const employeeName =
            document
                .getElementById("employeeName")
                .value
                .trim();

        const targetMonth =
            document
                .getElementById("targetMonth")
                .value;

        const resultArea =
            document.getElementById("result");

        const totalArea =
            document.getElementById("total");

        const calendarButton =
            document.getElementById("calendarButton");

        const calendarMessage =
            document.getElementById("calendarMessage");


        /*
         * 入力確認
         */
        if (employeeName === "") {
            alert("名前を入力してください");
            return;
        }

        if (targetMonth === "") {
            alert("対象年月を選択してください");
            return;
        }

        if (fileInput.files.length === 0) {
            alert("Excelファイルを選択してください");
            return;
        }


        /*
         * 前回の結果を初期化
         */
        resultArea.innerHTML = "";
        totalArea.textContent = "";
        calendarMessage.textContent = "";
        calendarButton.hidden = true;

        shiftEvents = [];


        /*
         * YYYY-MMを年と月に分ける
         */
        const [targetYear, targetMonthNumber] =
            targetMonth.split("-").map(Number);

        if (
            !Number.isInteger(targetYear) ||
            !Number.isInteger(targetMonthNumber) ||
            targetMonthNumber < 1 ||
            targetMonthNumber > 12
        ) {
            alert("対象年月が正しくありません");
            return;
        }


        const file = fileInput.files[0];

        try {
            const buffer =
                await file.arrayBuffer();

            const workbook =
                new ExcelJS.Workbook();

            await workbook.xlsx.load(buffer);

            let allWorkTime = 0;


            /*
             * シートを順番に確認
             */
            for (
                let sheetNum = 0;
                sheetNum < workbook.worksheets.length;
                sheetNum++
            ) {
                const worksheet =
                    workbook.worksheets[sheetNum];

                /*
                 * シート名から「○日」を取得
                 *
                 * 例:
                 * 1日       → 1
                 * １日      → 1
                 * 01日      → 1
                 * ０１日    → 1
                 * 11月1日   → 1
                 */
                const day =
                    extractDayFromSheetName(
                        worksheet.name
                    );

                /*
                 * 日付を取得できないシートは除外
                 */
                if (day === null) {
                    console.log(
                        "勤務表ではないためスキップ:",
                        worksheet.name
                    );

                    continue;
                }


                /*
                 * 入力された年月に実在する日付か確認
                 *
                 * 例:
                 * 2026年2月30日などは除外
                 */
                if (
                    !isValidDate(
                        targetYear,
                        targetMonthNumber,
                        day
                    )
                ) {
                    console.warn(
                        "対象年月には存在しない日付です:",
                        worksheet.name
                    );

                    continue;
                }


                /*
                 * 名前を探す
                 */
                for (let row = 1; row < 100; row++) {
                    const nameCell =
                        worksheet.getCell(`A${row}`);

                    const excelEmployeeName =
                        String(nameCell.value ?? "")
                            .trim();

                    if (
                        excelEmployeeName !==
                        employeeName
                    ) {
                        continue;
                    }


                    let startTimeCell = null;
                    let startColumn = null;
                    let workTime = 0;
                    let workCellCount = 0;


                    /*
                     * 勤務色が付いたセルを確認
                     */
                    for (
                        let column = 1;
                        column < 40;
                        column++
                    ) {
                        const cell =
                            worksheet.getCell(
                                row,
                                column
                            );

                        if (!isWorkCell(cell)) {
                            continue;
                        }

                        if (startTimeCell === null) {
                            startTimeCell =
                                worksheet.getCell(
                                    3,
                                    column
                                );

                            startColumn = column;
                        }

                        workTime += 0.5;
                        workCellCount++;
                    }


                    /*
                     * 勤務時間が見つかった場合
                     */
                    if (
                        workTime > 0 &&
                        startTimeCell !== null &&
                        startColumn !== null
                    ) {
                        const finishTimeCell =
                            worksheet.getCell(
                                3,
                                startColumn +
                                    workCellCount
                            );

                        const startTime =
                            getTime(startTimeCell);

                        const finishTime =
                            getTime(finishTimeCell);

                        /*
                         * 時刻をDateへ変換
                         */
                        const event =
                            createShiftEvent(
                                targetYear,
                                targetMonthNumber,
                                day,
                                startTime,
                                finishTime
                            );

                        if (event !== null) {
                            shiftEvents.push(event);
                        }


                        allWorkTime += workTime;


                        /*
                         * 結果を画面へ表示
                         */
                        const workElement =
                            document.createElement(
                                "div"
                            );

                        workElement.className = "work";

                        const dateElement =
                            document.createElement(
                                "strong"
                            );

                        dateElement.textContent =
                            `${targetYear}年` +
                            `${targetMonthNumber}月` +
                            `${day}日`;

                        const workTimeElement =
                            document.createElement(
                                "div"
                            );

                        workTimeElement.textContent =
                            `勤務時間：${workTime}時間`;

                        const timeElement =
                            document.createElement(
                                "div"
                            );

                        timeElement.textContent =
                            `${startTime} ～ ${finishTime}`;

                        workElement.appendChild(
                            dateElement
                        );

                        workElement.appendChild(
                            workTimeElement
                        );

                        workElement.appendChild(
                            timeElement
                        );

                        resultArea.appendChild(
                            workElement
                        );

                        /*
                         * 同じシートで対象者を
                         * 見つけたため行検索を終了
                         */
                        break;
                    }
                }
            }


            totalArea.textContent =
                `合計勤務時間：${allWorkTime}時間`;


            /*
             * シフトが1件以上あれば
             * カレンダー登録ボタンを表示
             */
            if (shiftEvents.length > 0) {
                shiftEvents.sort(
                    (eventA, eventB) =>
                        eventA.start -
                        eventB.start
                );

                calendarButton.hidden = false;

                calendarMessage.textContent =
                    `${shiftEvents.length}件の勤務を` +
                    "カレンダーへ登録できます。";
            } else {
                calendarButton.hidden = true;

                calendarMessage.textContent =
                    "登録できる勤務情報がありません。";
            }
        } catch (error) {
            console.error(error);

            resultArea.textContent =
                "Excelファイルを読み込めませんでした。" +
                ".xlsx形式のファイルを確認してください。";

            totalArea.textContent = "";
            calendarButton.hidden = true;
        }
    });


/*
 * カレンダー登録ボタン
 */
document
    .getElementById("calendarButton")
    .addEventListener("click", function () {
        if (shiftEvents.length === 0) {
            alert("登録できる勤務情報がありません");
            return;
        }

        downloadICS(shiftEvents);
    });


/*
 * 全角文字を半角文字へ正規化
 *
 * 例:
 * １日      → 1日
 * ０１日    → 01日
 * １１月１日 → 11月1日
 * 　１日　   → 1日
 */
function normalizeText(value) {
    return String(value ?? "")
        .normalize("NFKC")
        .replace(/\s+/g, "");
}


/*
 * シート名から日だけを取得
 *
 * 対応例:
 * 1日
 * １日
 * 01日
 * ０１日
 * 11月1日
 * １１月１日
 *
 * 「勤務表1日」のような名前にも対応します。
 */
function extractDayFromSheetName(sheetName) {
    const normalizedName =
        normalizeText(sheetName);

    const match =
        normalizedName.match(
            /(?:\d{1,2}月)?(\d{1,2})日/
        );

    if (match === null) {
        return null;
    }

    const day = Number(match[1]);

    if (
        !Number.isInteger(day) ||
        day < 1 ||
        day > 31
    ) {
        return null;
    }

    return day;
}


/*
 * 指定年月に日付が存在するか確認
 */
function isValidDate(year, month, day) {
    const date =
        new Date(
            year,
            month - 1,
            day
        );

    return (
        date.getFullYear() === year &&
        date.getMonth() === month - 1 &&
        date.getDate() === day
    );
}


/*
 * 勤務色セルの判定
 */
function isWorkCell(cell) {
    const workColors = [
        "FFFFFF00",
        "FFFFC000",
        "FFFF0000",
        "FF00B0F0",
        "FFFF99FF",
        "FF9933FF",
        "FF99FFCC"
    ];

    const foregroundColor =
        cell.fill?.fgColor;

    const argb =
        foregroundColor?.argb;

    const isWorkColor =
        workColors.includes(argb);

    const isThemeColor =
        foregroundColor?.theme === 0 &&
        foregroundColor?.tint ===
            -0.1499984740745262;

    return isWorkColor || isThemeColor;
}


/*
 * Excelセルから時刻を取得
 *
 * 戻り値:
 * 09:00
 * 17:30
 */
function getTime(timeCell) {
    const time = timeCell.value;

    /*
     * ExcelJSがDateとして返した場合
     */
    if (time instanceof Date) {
        const hour =
            String(time.getUTCHours())
                .padStart(2, "0");

        const minute =
            String(time.getUTCMinutes())
                .padStart(2, "0");

        return `${hour}:${minute}`;
    }


    /*
     * Excelの時刻が数値の場合
     *
     * 例:
     * 0.5 = 12:00
     */
    if (
        typeof time === "number" &&
        Number.isFinite(time)
    ) {
        const totalMinutes =
            Math.round(
                (time % 1) * 24 * 60
            );

        const hour =
            String(
                Math.floor(totalMinutes / 60) % 24
            ).padStart(2, "0");

        const minute =
            String(totalMinutes % 60)
                .padStart(2, "0");

        return `${hour}:${minute}`;
    }


    /*
     * 文字列の場合
     */
    if (typeof time === "string") {
        return normalizeTime(time);
    }

    return "";
}


/*
 * 時刻表記をHH:MM形式へ正規化
 *
 * 例:
 * 9:00     → 09:00
 * ９：００ → 09:00
 * 9時30分  → 09:30
 */
function normalizeTime(value) {
    const normalized =
        String(value ?? "")
            .normalize("NFKC")
            .trim()
            .replace(/：/g, ":");

    const match =
        normalized.match(
            /^(\d{1,2})(?::|時)(\d{1,2})?(?:分)?/
        );

    if (match === null) {
        return "";
    }

    const hour = Number(match[1]);
    const minute = Number(match[2] ?? 0);

    if (
        !Number.isInteger(hour) ||
        !Number.isInteger(minute) ||
        hour < 0 ||
        hour > 23 ||
        minute < 0 ||
        minute > 59
    ) {
        return "";
    }

    return (
        String(hour).padStart(2, "0") +
        ":" +
        String(minute).padStart(2, "0")
    );
}


/*
 * 勤務イベントを作成
 */
function createShiftEvent(
    year,
    month,
    day,
    startTime,
    finishTime
) {
    const startParts =
        parseTime(startTime);

    const finishParts =
        parseTime(finishTime);

    if (
        startParts === null ||
        finishParts === null
    ) {
        console.warn(
            "勤務時刻を認識できません:",
            startTime,
            finishTime
        );

        return null;
    }

    const start =
        new Date(
            year,
            month - 1,
            day,
            startParts.hour,
            startParts.minute,
            0
        );

    const end =
        new Date(
            year,
            month - 1,
            day,
            finishParts.hour,
            finishParts.minute,
            0
        );


    /*
     * 終了時刻が開始時刻以前なら、
     * 翌日までの勤務として扱う
     *
     * 例:
     * 22:00～02:00
     */
    if (end <= start) {
        end.setDate(end.getDate() + 1);
    }

    return {
        start,
        end,
        title: "温泉勤務"
    };
}


/*
 * HH:MMを数値へ変換
 */
function parseTime(value) {
    const normalized =
        normalizeTime(value);

    if (normalized === "") {
        return null;
    }

    const [hour, minute] =
        normalized.split(":").map(Number);

    return {
        hour,
        minute
    };
}


/*
 * ICSファイルを生成
 */
function createICS(events) {
    const lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Shift Viewer//Onsen Shift//JP",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:寿の湯シフト"
    ];

    events.forEach(function (event) {
        lines.push(
            "BEGIN:VEVENT",
            `UID:${createEventUID(event)}`,
            `DTSTAMP:${formatICSUtcDate(new Date())}`,
            `DTSTART;TZID=Asia/Tokyo:${formatICSLocalDate(event.start)}`,
            `DTEND;TZID=Asia/Tokyo:${formatICSLocalDate(event.end)}`,
            `SUMMARY:${escapeICSText(event.title)}`,
            "STATUS:CONFIRMED",
            "TRANSP:OPAQUE",
            "END:VEVENT"
        );
    });

    lines.push("END:VCALENDAR");

    /*
     * ICSではCRLF改行が推奨
     */
    return lines.join("\r\n");
}


/*
 * ICSのローカル日時
 *
 * 例:
 * 20261101T090000
 */
function formatICSLocalDate(date) {
    const year =
        date.getFullYear();

    const month =
        String(date.getMonth() + 1)
            .padStart(2, "0");

    const day =
        String(date.getDate())
            .padStart(2, "0");

    const hour =
        String(date.getHours())
            .padStart(2, "0");

    const minute =
        String(date.getMinutes())
            .padStart(2, "0");

    const second =
        String(date.getSeconds())
            .padStart(2, "0");

    return (
        `${year}${month}${day}` +
        `T${hour}${minute}${second}`
    );
}


/*
 * DTSTAMP用のUTC日時
 *
 * 例:
 * 20261101T000000Z
 */
function formatICSUtcDate(date) {
    const year =
        date.getUTCFullYear();

    const month =
        String(date.getUTCMonth() + 1)
            .padStart(2, "0");

    const day =
        String(date.getUTCDate())
            .padStart(2, "0");

    const hour =
        String(date.getUTCHours())
            .padStart(2, "0");

    const minute =
        String(date.getUTCMinutes())
            .padStart(2, "0");

    const second =
        String(date.getUTCSeconds())
            .padStart(2, "0");

    return (
        `${year}${month}${day}` +
        `T${hour}${minute}${second}Z`
    );
}


/*
 * ICS内の特殊文字をエスケープ
 */
function escapeICSText(value) {
    return String(value ?? "")
        .replace(/\\/g, "\\\\")
        .replace(/\r?\n/g, "\\n")
        .replace(/,/g, "\\,")
        .replace(/;/g, "\\;");
}


/*
 * イベントごとに安定したUIDを作る
 *
 * 同じ勤務情報からは同じUIDを生成するため、
 * カレンダーアプリが重複を判定しやすくなります。
 */
function createEventUID(event) {
    const startText =
        formatICSLocalDate(event.start);

    const endText =
        formatICSLocalDate(event.end);

    return (
        `onsen-shift-${startText}-${endText}` +
        "@shift-viewer.local"
    );
}


/*
 * ICSをダウンロード
 */
function downloadICS(events) {
    const icsContent =
        createICS(events);

    /*
     * UTF-8のBOMを付け、
     * 日本語の文字化けを防ぐ
     */
    const blob =
        new Blob(
            ["\uFEFF", icsContent],
            {
                type:
                    "text/calendar;charset=utf-8"
            }
        );

    const url =
        URL.createObjectURL(blob);

    const firstEvent =
        events[0];

    const fileYear =
        firstEvent.start.getFullYear();

    const fileMonth =
        String(
            firstEvent.start.getMonth() + 1
        ).padStart(2, "0");

    const link =
        document.createElement("a");

    link.href = url;
    link.download =
        `寿シフト-${fileYear}-${fileMonth}.ics`;

    document.body.appendChild(link);
    link.click();
    link.remove();


    /*
     * ダウンロード開始直後ではなく、
     * 少し待ってからメモリ上のURLを破棄
     */
    window.setTimeout(function () {
        URL.revokeObjectURL(url);
    }, 1000);
}