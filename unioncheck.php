<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>유니온 멤버 딜량 체크</title>
    <style>
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }

        th, td {
            border: 1px solid black;
            padding: 10px;
            text-align: left;
        }

        input[type="text"] {
            margin-bottom: 10px;
            padding: 8px;
        }
    </style>
</head>

<body>
    <div>
        <label for="nicknameInput">닉네임:</label>
        <input type="text" id="nicknameInput">
        <button onclick="addRow()">입력</button>
        <br>
        <label for="damageInput">딜량:</label>
        <input type="text" id="damageInput">
        <br>
        <br>
        <br>
    </div>
    <table id="data-table">
        <tr>
            <th>닉네임</th>
            <th>참여 횟수</th>
            <th>누적 딜량</th>
            <th></th>
        </tr>
    </table>


    <div>
        <label for="nonParticipantName">레이드 미참여자:</label>
        <input type="text" id="nonParticipantName">
        <label for="nonParticipantReason">사유:</label>
        <input type="text" id="nonParticipantReason">
        <button onclick="addNonParticipantRow()">추가</button>
        <br><br>
    </div>

    <div>
            <label for="totalMembersInput">전체 멤버수:</label>
            <input type="text" id="totalMembersInput">
            <button onclick="updateStats()">입력</button>
    </div>
    <div>
        <br>
        <label>참여자 수: <span id="participantCount">0</span></label><br>
        <label>참여율: <span id="participationRate">0%</span></label>
    </div>



    <div>
        <br><br>
        <label for="daySelector">일차 선택:</label>
        <select id="daySelector">
        <option value="1">1일차</option>
        <option value="2">2일차</option>
        <option value="3">3일차</option>
        <option value="4">4일차</option>
        <option value="5">5일차</option>
        <option value="6">6일차</option>
        <option value="7">7일차</option>    
        </select>
        <button onclick="exportToExcel()">저장</button>
    </div>


    


    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.0/dist/xlsx.full.min.js"></script>

    <script>
        function addRow() {
            var nickname = document.getElementById("nicknameInput").value;
            var damage = parseInt(document.getElementById("damageInput").value);

            var table = document.getElementById("data-table");
            var rows = table.getElementsByTagName("tr");

            for (var i = 1; i < rows.length; i++) {
                var currentNickname = rows[i].getElementsByTagName("td")[0].textContent;
                var currentDamage = parseInt(rows[i].getElementsByTagName("td")[2].textContent);

                if (nickname === currentNickname) {
                    // 이미 존재하는 닉네임인 경우 참여 횟수와 딜량 누적
                    var currentCount = parseInt(rows[i].getElementsByTagName("td")[1].textContent);
                    if (currentCount >= 3) {
                        alert('하루에 최대 3번까지만 유니온 레이드에 참여할 수 있습니다!');
                        return;
                    } else {
                        rows[i].getElementsByTagName("td")[1].textContent = currentCount + 1;
                        rows[i].getElementsByTagName("td")[2].textContent = currentDamage + damage;
                        sortTableByColumn(1);
                        return;
                    }
                }
            }

            // 새로운 행 추가
            var newRow = table.insertRow(-1);
            var cell1 = newRow.insertCell(0);
            var cell2 = newRow.insertCell(1);
            var cell3 = newRow.insertCell(2);
            cell1.textContent = nickname;
            cell2.textContent = 1;
            cell3.textContent = damage;
            


            var deleteButton = document.createElement("button");
            deleteButton.textContent = "삭제";
            deleteButton.onclick = function () {
                var row = this.parentNode.parentNode;
                row.parentNode.removeChild(row);
            };

            var cell4 = newRow.insertCell(3);
            cell4.appendChild(deleteButton);
        }


        function updateStats() {
            var totalMembers = parseInt(document.getElementById("totalMembersInput").value);
            var table = document.getElementById("data-table");
            var rows = table.getElementsByTagName("tr");
            var participantCount = 0;

            // 유효한 참여자만을 세어 참여 인원 계산
            for (var i = 1; i < rows.length; i++) {
                var currentCount = parseInt(rows[i].getElementsByTagName("td")[1].textContent);
                if (currentCount > 0) {
                    participantCount++;
                }
            }

            var participationRate = (participantCount / totalMembers) * 100;
            document.getElementById("participantCount").textContent = participantCount;
            document.getElementById("participationRate").textContent = participationRate.toFixed(2) + "%";
        }

        function addNonParticipantRow() {
            var nonParticipantName = document.getElementById("nonParticipantName").value;
            var nonParticipantReason = document.getElementById("nonParticipantReason").value;

            var table = document.getElementById("data-table");

            // 중복된 닉네임을 확인
            var rows = table.getElementsByTagName("tr");
            for (var i = 1; i < rows.length; i++) {
                var currentNickname = rows[i].getElementsByTagName("td")[0].textContent;
                if (nonParticipantName === currentNickname) {
                    alert('닉네임은 중복될 수 없습니다!');
                    return;
                }
            }

            // 중복이 없는 경우 행 추가
            var newRow = table.insertRow(-1); // 마지막에 행 추가
            var cell1 = newRow.insertCell(0);
            var cell2 = newRow.insertCell(1);
            var cell3 = newRow.insertCell(2);

            // 새로운 행에 레이드 미참여자의 정보 추가
            cell1.textContent = nonParticipantName;
            cell2.textContent = 0; // 참여 횟수를 0으로 설정
            cell3.textContent = nonParticipantReason; // 누적 딜량에 사유 추가

            

            // 삭제 버튼 추가
            var deleteButton = document.createElement("button");
            deleteButton.textContent = "삭제";
            deleteButton.onclick = function () {
                var row = this.parentNode.parentNode;
                row.parentNode.removeChild(row);
            };

            var cell4 = newRow.insertCell(3);
            cell4.appendChild(deleteButton);

            sortTableByColumn(1);
        }


        function exportToExcel() {
            var selectedDay = document.getElementById("daySelector").value;
            var filename = selectedDay + "일차_유니온_데이터.xlsx";

            var table = document.getElementById("data-table");
            var workbook = XLSX.utils.table_to_book(table);

            // 엑셀 데이터에 표 외의 요소들 추가
            var sheet = workbook.Sheets[workbook.SheetNames[0]]; // 첫 번째 시트 선택
            sheet["A7"] = { t: "s", v: "전체 멤버수", s: { font: { bold: true } } };
            sheet["B7"] = { t: "n", v: parseInt(document.getElementById("totalMembersInput").value) };

            // 레이드 미참여자 정보 추가
            var nonParticipantName = document.getElementById("nonParticipantName").value;
            var nonParticipantReason = document.getElementById("nonParticipantReason").value;
            sheet["A8"] = { t: "s", v: "레이드 미참여자", s: { font: { bold: true } } };
            sheet["B8"] = { t: "s", v: nonParticipantName + " (" + nonParticipantReason + ")" };

            // 참여자 수와 참여율 정보 추가
            sheet["A9"] = { t: "s", v: "참여자 수", s: { font: { bold: true } } };
            sheet["B9"] = { t: "n", v: parseInt(document.getElementById("participantCount").textContent) };
            sheet["A10"] = { t: "s", v: "참여율", s: { font: { bold: true } } };
            sheet["B10"] = { t: "s", v: document.getElementById("participationRate").textContent };

            // 엑셀 파일 생성
            XLSX.writeFile(workbook, filename);
        }

        function sortTableByColumn(columnIndex) {
            var table, rows, switching, i, x, y, shouldSwitch;
            table = document.getElementById("data-table");
            switching = true;
            while (switching) {
                switching = false;
                rows = table.getElementsByTagName("tr");
                for (i = 1; i < rows.length - 1; i++) {
                    shouldSwitch = false;
                    x = parseInt(rows[i].getElementsByTagName("td")[columnIndex].textContent);
                    y = parseInt(rows[i + 1].getElementsByTagName("td")[columnIndex].textContent);

                    // 참여 횟수가 같을 경우 딜량 차이로 정렬
                    if (x === y) {
                        var damageX = parseInt(rows[i].getElementsByTagName("td")[2].textContent);
                        var damageY = parseInt(rows[i + 1].getElementsByTagName("td")[2].textContent);
                        if (damageX < damageY) {
                            shouldSwitch = true;
                            break;
                        }
                    } else {
                        if (x < y) {
                            shouldSwitch = true;
                            break;
                        }
                    }
                }
                if (shouldSwitch) {
                    rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                    switching = true;
                }
            }
        }



    </script>


</body>

</html>
