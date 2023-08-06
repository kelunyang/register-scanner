<template>
  <el-dialog
    v-model="dialogs[0].show"
    :fullscreen="true"
    :show-close="false"
    :title="dialogs[0].title">
    <el-space direction="vertical" fill wrap style="width: 100%">
      <el-steps :active="stepIndicator" finish-status="finish" align-center>
        <el-step :title="step.title" v-for="(step, index) in dialogs" :key="index" :status="step.status" />
      </el-steps>
      <el-alert :title="'發生錯誤'+sysMsg.title" :type="sysMsg.type" show-icon v-if="sysMsg.status">
        <template #default>
          <span style="font-size: 1.5em">
            {{ sysMsg.message }}
          </span>
        </template>
      </el-alert>
      <div class="xs12" style="font-size: 1em; color: #666; text-align: center;" v-if="sheetList.length === 0">無資料</div>
      <el-table :data="sheetList" stripe style="width: 100%" v-else>
        <el-table-column prop="dueDate" label="表單名稱" sortable>
          <template #default="scope">
            <el-tag
              v-for="tag in tagFinder(scope.row.title)"
              :key="tag.id"
              :color="tag.color"
              style="margin:1px"
              effect="dark"
            >
              {{ tag.name }}
            </el-tag>
            <span style="font-weight: bold">{{ removeTag(scope.row.title) }}</span><br/>
            <span>填寫至：{{dateConverter(scope.row.dueDate) }}</span><br/>
          </template>
        </el-table-column>
        <el-table-column label="">
          <template #default="scope">
            <div class="buttonBlock">
            <el-button class="ma1 pa2" size="large" type="primary" v-on:click="openSheet(scope.row.id)">輸入表單</el-button>
            </div>
          </template>
        </el-table-column>
      </el-table>
      <div class="footerText">Developer: <a class="cleanLink" href="mailto:kelunyang@outlook.com">Kelunyang</a>@LKSH 2023 <a style="color:#CCC" target="_blank" href="https://github.com/kelunyang/register-scanner" >GITHUB</a></div>
    </el-space>
  </el-dialog>
  <el-dialog
    :show-close="false"
    v-model="dialogs[1].show"
    :fullscreen="true"
    :title="dialogs[1].title + '問卷：' + currentQuery.title">
    <el-space direction="vertical" fill wrap style="width: 100%">
      <el-steps :active="stepIndicator" finish-status="finish" align-center>
        <el-step :title="step.title" v-for="(step, index) in dialogs" :key="index" :status="step.status" />
      </el-steps>
      <el-alert :title="sysMsg.title" :type="sysMsg.type" show-icon v-if="sysMsg.status">
        <template #default>
          <span style="font-size: 1.5em">
            {{ sysMsg.message }}
          </span>
        </template>
      </el-alert>
      <el-space direction="vertical" fill wrap class="ma1 pa2 xs12">
        <div class="qTitle xs12">帳號</div>
        <el-input
          size="large"
          class="xs12"
          label="帳號"
          v-model="currentQuery.account"
          outline>
        </el-input>
      </el-space>
      <el-space direction="vertical" fill wrap class="ma1 pa2 xs12">
        <div class="qTitle xs12">密碼</div>
        <el-input
          size="large"
          class="xs12"
          label="密碼"
          v-model="currentQuery.password"
          :show-password="true"
          outline>
        </el-input>
      </el-space>
      <el-button v-if="!connecting" class="ma1 pa1 xs12" size="large" type="danger" :disabled="checkAuth" v-show="!connecting" v-on:click="checkLogin()">{{ checkAuth ? "請輸入帳號密碼後登入" : "登入表單" }}</el-button>
      <el-button class="ma1 pa2 xs12" size="large" type="primary" v-on:click="reloadPage()" v-if="!loginStatus">回到問卷列表</el-button>
      <div class="footerText">Developer: <a class="cleanLink" href="mailto:kelunyang@outlook.com">Kelunyang</a>@LKSH 2023 <a style="color:#CCC" target="_blank" href="https://github.com/kelunyang/register-scanner" >GITHUB</a></div>
    </el-space>
  </el-dialog>
  <el-dialog
    v-model="dialogs[2].show"
    :fullscreen="true"
    :show-close="false"
    :title="dialogs[2].title + '：' + currentQuery.title + '／' + currentQuery.account">
      <el-space direction="vertical" fill wrap style="width: 100%">
        <el-steps :active="stepIndicator" finish-status="finish" align-center>
          <el-step :title="step.title" v-for="(step, index) in dialogs" :key="index" :status="step.status" />
        </el-steps>
        <el-space v-if="debugPrecentage > 0" direction="horizonal" fill wrap style="width: 100%">
          <span>測試模式倒數計時器</span>
          <el-progress :percentage="debugPrecentage" :color="progressColors"/>
        </el-space>
        <el-alert :title="sysMsg.title" :type="sysMsg.type" show-icon v-if="sysMsg.status">
          <template #default>
            <span style="font-size: 1.5em">
              {{ sysMsg.message }}
            </span>
          </template>
        </el-alert>
        <el-space direction="vertical" fill wrap class="ma1 pa2 xs12 breakword" v-for="column in currentQuery.inputColumns" :key="column.id" :style="formatDetector('P', column) ? { background: '#ffffcc' } : {}">
          <div class="qTitle xs12">{{ column.name }}</div>
          <el-input
            v-if="formatDetector('P', column)"
            size="large"
            class="xs12"
            :label="column.name"
            v-model="column.value"
            @keyup.enter="sendInsert()"
            outline>
          </el-input>
          <el-input
            v-if="formatDetector('T', column)"
            size="large"
            class="xs12"
            :label="column.name"
            v-model="column.value"
            v-on:change="valField(column)"
            outline>
          </el-input>
          <el-select
            v-if="formatDetector('S', column)"
            v-model="column.value"
            class="xs12"
            :placeholder="column.name"
            v-on:change="valField(column)"
            size="large">
            <el-option
              v-for="item in column.format.split(';')"
              :key="item+column.id+'key'"
              :label="item"
              :value="item"
            />
          </el-select>
          <div class="alertWord" v-if="column.status !== ''">{{ column.status }}</div>
          <div class="captionWord" v-if="column.status === ''">{{ formatHelper(column) }}</div>
        </el-space>
        <el-button v-if="!connecting" class="ma1 pa2 xs12" size="large" type="danger" v-on:click="sendInsert()" :disabled="!checkInsert">
          {{ checkInsert ? "送出" : "你的格式有誤，請檢查" }}
        </el-button>
        <el-button class="ma1 pa2 xs12" size="large" type="success" v-on:click="startSelection()">設定回傳欄位({{ currentQuery.selectedViewable.length }})</el-button>
        <el-button v-if="currentQuery.debugData.length > 0" class="ma1 pa2 xs12" size="large" type="primary" v-on:click="showDebug()">測試寫入數據</el-button>
        <el-button v-if="currentQuery.debugData.length > 0" class="ma1 pa2 xs12" size="large" type="primary" v-on:click="downloadDebug()">下載寫入時間數據</el-button>
        <el-button class="ma1 pa2 xs12" size="large" type="primary" v-on:click="reloadPage()" v-if="!loginStatus">回到問卷列表</el-button>
      </el-space>
  </el-dialog>
  <el-drawer
    v-model="viewableW"
    title="設定登記成功後的回傳確認欄位"
    direction="btt"
    show-close="false"
    size="90%"
  >
    <el-space direction="vertical" fill wrap style="width: 100%">
      <el-transfer
        class="ma1 pa2 xs12"
        v-model="currentQuery.selectedViewable"
        filterable
        :filter-method="filterMethod"
        filter-placeholder="在此可以打字搜尋"
        :data="currentQuery.viewableColumns"
        v-on:right-check-change="chooseSelection"
        target-order="push"
        :titles="['候選名單', '已選名單']"
        :button-texts="['移出已選', '移入已選']"
      >
      </el-transfer>
      <el-space direction="horizonal" fill wrap class="ma1 pa2 xs12" v-if="currentQuery.modified.length > 0">
        <el-button class="ma1 pa2 xs12" size="large" type="success" @click="selectionMove(0)">將已選的{{ currentQuery.modified.length }}個選項置頂</el-button>
        <el-button class="ma1 pa2 xs12" size="large" type="success" @click="selectionMove(2)">將已選的{{ currentQuery.modified.length }}個選項向上一格</el-button>
        <el-button class="ma1 pa2 xs12" size="large" type="success" @click="selectionMove(3)">將已選的{{ currentQuery.modified.length }}個選項向下一格</el-button>
        <el-button class="ma1 pa2 xs12" size="large" type="success" @click="selectionMove(1)">將已選的{{ currentQuery.modified.length }}個選項置底</el-button>
      </el-space>
      <el-button class="ma1 pa2 xs12" size="large" type="danger" v-on:click="endSelection()">選擇完畢！</el-button>
    </el-space>
  </el-drawer>
  <el-drawer
    v-model="debugW"
    title="測試寫入數據"
    direction="btt"
    show-close="false"
    size="90%"
  >
    <el-space direction="vertical" fill wrap style="width: 100%">
      <el-alert title="測試狀態" type="info" show-icon>
        <template #default>
          <span style="font-size: 1.5em">
            測試已{{ debugStat ? "啟動" : "關閉" }}，你的測試資料有{{ currentQuery.debugData.length }}筆，第一筆資料預覽：{{ previewData }}
          </span>
        </template>
      </el-alert>
      <el-space direction="horizonal" fill wrap style="width: 100%">
        <span>執行間隔時間（秒）：目前設定為{{ debugInterval }}秒</span>
        <el-slider v-model="debugInterval" :step="10" :min="10" :max="60" />
      </el-space>
      <el-space direction="horizonal" fill wrap style="width: 100%">
        <span>冷卻混淆時間（秒）：目前設定為{{ sleepTime }}秒</span>
        <el-slider v-model="sleepTime" :step="1" :min="0" :max="debugInterval" />
      </el-space>
      <el-space direction="horizonal" fill wrap style="width: 100%">
        <span>執行區間（分）：目前設定為{{ debugDuration }}分</span>
        <el-slider v-model="debugDuration" :step="1" :min="5" :max="180" />
      </el-space>
      <el-switch
        v-model="sufferSelect"
        class="mb-2"
        active-text="針對選擇型欄位亂數選擇"
        inactive-text="選擇型欄位按照原始資料表填入（沒有就留空）"
      />
      <el-button class="ma1 pa2 xs12" size="large" type="danger" v-on:click="exeDebug()">{{ debugStat ? "停止" : "啟動" }}測試</el-button>
      <el-button class="ma1 pa2 xs12" size="large" type="primary" v-on:click="hideDebug()">關閉測試設定畫面</el-button>
    </el-space>
  </el-drawer>
</template>
  
<script>
  import { nextTick, reactive, onMounted, ref, computed  } from 'vue';
  import { ElMessage } from 'element-plus';
  import dayjs from 'dayjs';
  import randomColor from 'randomcolor';
  import { v4 as uuidv4 } from 'uuid';
  import _ from'lodash';

  export default {
    setup() {
      const progressColors = [
        { color: '#ffffcc', percentage: 20 },
        { color: '#ffcc99', percentage: 40 },
        { color: '#ff6666', percentage: 60 },
        { color: '#ff3333', percentage: 80 },
        { color: '#990000', percentage: 100 },
      ]
      let debugPrecentage = ref(0);
      let sufferSelect = ref(false);
      let sleepTime = ref(0);
      let debugTimer = null;
      let debugStat = ref(false);
      let debugInterval = ref(10);
      let debugDuration = ref(10);
      let debugW = ref(false)
      let viewableW = ref(false);
      let colors = [];
      let debugDB = [];
      let connecting = ref(false);
      let sheetList = ref([]);
      let dialogs = ref([{
        title: "表單列表",
        show: true,
        status: "wait"
      },{
        title: "登入表單",
        show: false,
        status: "wait"
      },{
        title: "表單輸入",
        show: false,
        status: "wait"
      }]);
      let stepIndicator = ref(0);
      const currentQuery = reactive({
        id: "",
        successWord: "",
        failWord: "",
        duplicateWord: "",
        title: "",
        count: 0,
        account: "",
        password: "",
        viewableColumns: [],
        inputColumns: [],
        selectedViewable: [],
        modified: [],
        debugData: [],
        mode: ""
      });
      const sysMsg = reactive({
        status: false,
        type: "error",
        message: "",
        title: ""
      });
      const formatDetector = (formatReg, column) => {
        if((new RegExp(formatReg)).test(column.type)) {
          return true;
        }
        return false;
      }
      const checkAuth = computed(() => {
        if(currentQuery.account === "" || currentQuery.password === "") {
          return true;
        }
        return false;
      })
      const dateConverter = (tick) => {
        if(tick === "" || tick === undefined) {
          return "無"
        } else {
          let dayObj = dayjs(tick);
          return dayObj.format('YYYY-MM-DD HH:mm:ss')
        }
      }
      const changeStep = (name, currentStatus, previousStatus, nextStatus) => {
        for(let i=0; i< dialogs.value.length; i++) {
          if(dialogs.value[i].title === name) {
            dialogs.value[i].status = currentStatus;
            stepIndicator = i;
            if(dialogs.value[i].status === "wait") {
              dialogs.value[i].show = false;
            } else {
              dialogs.value[i].show = true;
            }
          }
        }
        for(let i=0; i<dialogs.value.length; i++) {
          if(i < stepIndicator) {
            dialogs.value[i].status = previousStatus;
          }
          if(i > stepIndicator) {
            dialogs.value[i].status = nextStatus;
          }
        }
        resetMsg();
      };
      const selectionMove = (direction) => {
        let currentIndex = -1;
        let foundIndexs = [];
        let newIndex = 0;
        if(direction === 0) {
          newIndex = 0;
        }
        for(let i=0; i<currentQuery.modified.length; i++) {
          let nowIndex = _.findIndex(currentQuery.selectedViewable, (item) => {
            return item === currentQuery.modified[i];
          });
          if(nowIndex > -1) {
            foundIndexs.push(nowIndex);
          }
        }
        foundIndexs = foundIndexs.sort();
        if(direction === 2) {
          currentIndex = foundIndexs[0];
        } else {
          currentIndex = foundIndexs[foundIndexs.length - 1];
        }
        if(currentIndex !== -1) {
          let tempSelected = [...currentQuery.selectedViewable];
          currentQuery.selectedViewable.splice(0);
          for(let i=0; i<tempSelected.length; i++) {
            let selected = _.filter(foundIndexs, (item) => {
              return item === i;
            });
            if(selected.length === 0) {
              currentQuery.selectedViewable.push(tempSelected[i]);
            }
          }
          if(direction === 2) {
            newIndex = currentIndex - 1 > 0 ? currentIndex - 1 : 0;
          } else if(direction === 3) {
            newIndex = currentIndex + 1 > currentQuery.selectedViewable.length ? currentQuery.selectedViewable.length : currentIndex + 1;
          } else if(direction === 1) {
            newIndex = currentQuery.selectedViewable.length > 0 ? currentQuery.selectedViewable.length : 0;
          }
          for(let i=0; i<foundIndexs.length; i++) {
            currentQuery.selectedViewable.splice(newIndex, 0, tempSelected[foundIndexs[i]]);
            newIndex++;
          }
        }
      }
      const removeTag = (title) => {
        return (title.match(/(?:.(?!\S*\]))+/))[0].replace(/\]/,"");
      }
      const tagFinder = (title) => {
        let tags = title.match(/\[[^\]]+\]/g);
        let tagResult = [];
        if(tags !== null) {
          for(let k=0; k<tags.length; k++) {
            tagResult.push({
              name: tags[k].replace(/\[|\]/g,""),
              color: colors[ k % colors.length ],
              id: uuidv4()
            });
          }
        }
        return tagResult;
      }
      const reloadPage = () => {
        google.script.run
          .withSuccessHandler(function(url){
            window.open(url,'_top');
          })
          .getScriptURL();
      };
      const checkLogin = () => {
        ElMessage("檢查登入中");
        connecting.value = true;
        google.script.run.withSuccessHandler((queryObj) => {
          if(!queryObj.status) {
            changeStep("登入表單", "error", "wait", "wait");
            resetMsg();
            sysMsg.type = "error";
            sysMsg.message = "帳號密碼驗證失敗，請檢查帳號密碼";
            sysMsg.title = "帳號密碼失敗";
            sysMsg.status = true;
          } else {
            currentQuery.mode = queryObj.result.mode;
            currentQuery.debugData = queryObj.result.debugData;
            currentQuery.viewableColumns = queryObj.result.viewableColumns;
            currentQuery.inputColumns = queryObj.result.inputColumns;
            currentQuery.selectedViewable = localStorage.getItem(currentQuery.id) === undefined || localStorage.getItem(currentQuery.id) === null ? [] : JSON.parse(localStorage.getItem(currentQuery.id));
            changeStep("表單輸入", "process", "wait", "wait");
          }
          connecting.value = false;
        })
        .withFailureHandler((data) => {
          resetMsg();
          sysMsg.type = "error";
          sysMsg.message = data;
          sysMsg.title = "主機端發生錯誤";
          sysMsg.status = true;
          connecting.value = false;
        })
        .checkAuth(currentQuery.account, currentQuery.password, currentQuery.id);
      }
      const formatHelper = (column) => {
          let tip = "";
          if(formatDetector('T', column)) {
            if(column.format === "") {
              tip = "文字";
            } else {
              let regexConfig = column.format.split("::");
              tip = regexConfig[0];
            }
          } else if(formatDetector('S', column)) {
            tip = "請從選單中選一個正確的值";
          } else if(formatDetector('P', column)) {
            tip = "本欄為主鍵，必須輸入，輸入完成按Enter或是掃描條碼會立刻送出"
          }
          let must = column.must ? "[必填]" : "";
          return must + "格式：" + tip + "[輸入後點其他區域會重新檢查本欄位格式]";
      }
      const checkInsert = computed(() => {
        return _.every(currentQuery.inputColumns, (column) => {
          return column.status === "";
        })
      });
      const sendInsert = () => {
        for(let i=0; i<currentQuery.inputColumns.length; i++) {
          valField(currentQuery.inputColumns[i]);
        }
        if(_.every(currentQuery.inputColumns, (column) => {
          return column.status === ""
        })) {
          ElMessage("表單送出中...");
          let PColumn = _.find(currentQuery.inputColumns, (column) => {
            return column.type === "P"
          });
          let transactionStart = new Date().getTime();
          connecting.value = true;
          google.script.run.withSuccessHandler((result) => {
            if(result.status) {
              currentQuery.count++;
              sysMsg.title = "送出成功";
              sysMsg.type = "info";
              sysMsg.status = true;
              sysMsg.message = currentQuery.successWord + "／主機回傳：" + _.join(result.result, ",") + "，你登入後已完成輸入" + currentQuery.count + "筆資料";
              if(debugStat.value) {
                sysMsg.message += "（測試模式）"
              }
              if(currentQuery.debugStat) {
                debugDB.push({
                  ID: PColumn.value,
                  startTime: transactionStart,
                  endTime: new Date().getTime(),
                  result: "成功寫入"
                });
              }
              for(let i=0;i<currentQuery.inputColumns.length; i++) {
                currentQuery.inputColumns[i].value = "";
              }
            } else {
              sysMsg.title = "寫入主機失敗";
              sysMsg.type = "error";
              sysMsg.status = true;
              sysMsg.message = "主機回傳：" + result.result;
              if(currentQuery.debugStat) {
                debugDB.push({
                  ID: PColumn.value,
                  startTime: transactionStart,
                  endTime: new Date().getTime(),
                  result: "寫入失敗"
                });
              }
            }
            connecting.value = false;
          })
          .withFailureHandler((data) => {
            sysMsg.title = "主機發生錯誤";
            sysMsg.type = "error";
            sysMsg.status = true;
            sysMsg.message = "主機回傳：" + data.message;
            if(currentQuery.debugStat) {
              debugDB.push({
                ID: PColumn.value,
                startTime: transactionStart,
                endTime: new Date().getTime(),
                result: "遠端錯誤"
              });
            }
          })
          .writeRecord(currentQuery.account, currentQuery.password, currentQuery.id, currentQuery.inputColumns, currentQuery.selectedViewable);
        } else {
          ElMessage("格式預檢錯誤，無法送出表單");
          sysMsg.title = "格式預檢錯誤";
          sysMsg.type = "error";
          sysMsg.status = true;
          sysMsg.message = "格式檢查失敗，你無法送出表單，請檢查格式（每個欄位的紅字處）";
        }
      }
      const chooseSelection = (selected) => {
        currentQuery.modified = selected;
        let checkSelected = _.find(currentQuery.selectedViewable, (selected) => {
          return selected > currentQuery.viewableColumns.length;
        })
        if(checkSelected > -1) {
          currentQuery.selectedViewable = localStorage.getItem(currentQuery.id) === undefined || localStorage.getItem(currentQuery.id) === null ? [] : JSON.parse(localStorage.getItem(currentQuery.id));
          ElMessage("你怎麼選到選項裡沒有的值的？已選欄位已被清空");
        }
      }
      const valField = (column) => {
        column.status = "";
        if(column.must) { //先檢查是否為空
          if(column.value === "") {
            column.status = "這個欄位必需有值！";
          } else {
            column.status = "";
          }
        }
        let formatCheck = column.must;
        if(!formatCheck) {
          formatCheck = column.value !== "";
        }
        if(/借用/.test(currentQuery.mode)) {
          formatCheck = column.state;
        }
        if(column.status === "") {  //最後檢查格式
          if(formatCheck) {
            if(formatDetector('T', column)) {
              if(column.format !== "") {
                let regexConfig = column.format.split("::");
                if(new RegExp(regexConfig[1]).test(column.value)) {
                  column.value = column.value.replace(/台(北|中|南|灣)/,'臺$1');
                  column.status = "";
                } else {
                  column.status = "格式提示為「" + regexConfig[0] + "」";
                }
              }
            } else if(formatDetector('S', column)) {
              let selection = column.format.split(";");
              if(selection.length > 0) {
                if(_.includes(selection, column.value)) {
                  column.status = "";
                } else {
                  column.status = "你真的是用選單選出來的值嗎？（還是你沒選）";
                }
              } else {
                column.status = "這個選單沒有提供選項？快去連絡管理員吧";
              }
            }
          }
        }
      }
      const openSheet = (id) => {
        let selectedSheet = _.filter(sheetList.value, (sheet) => {
          return sheet.id === id;
        })
        if(selectedSheet.length > 0) {
          currentQuery.title = selectedSheet[0].title;
          currentQuery.id = selectedSheet[0].id;
          currentQuery.successWord = selectedSheet[0].successWord;
          currentQuery.duplicateWord = selectedSheet[0].duplicateWord;
          currentQuery.failWord = selectedSheet[0].failWord;
          currentQuery.count = 0;
          changeStep("登入表單", "process", "wait", "wait");
        }
      }
      const endSelection = () => {
        currentQuery.modified = [];
        let checkSelected = _.find(currentQuery.selectedViewable, (selected) => {
          return selected >= currentQuery.viewableColumns.length;
        });
        if(checkSelected > -1) {
          currentQuery.selectedViewable = localStorage.getItem(currentQuery.id) === undefined || localStorage.getItem(currentQuery.id) === null ? [] : JSON.parse(localStorage.getItem(currentQuery.id));
          ElMessage("你怎麼選到不再欄位清單裡的欄位？已選欄位已被清空，請重新選擇");
        } else {
          localStorage.setItem(currentQuery.id, JSON.stringify(currentQuery.selectedViewable));
          ElMessage("回傳欄位已儲存！下次在這個瀏覽器開啟就不用再設定囉");
          viewableW.value = false;
        }
        resetMsg();
      }
      const startSelection = () => {
        viewableW.value = true;
        currentQuery.modified = [];
        resetMsg();
      }
      const resetMsg = () => {
        sysMsg.title = "";
        sysMsg.type = "";
        sysMsg.status = false;
        sysMsg.message = "";
      }
      const downloadDebug = () => {
        let output = "\ufeff"+ Papa.unparse(debugDB) + "\r\n本資料產生時間," + dayjs().format('YYYY-MM-DD HH:mm:ss');
        let blob = new Blob([output], { type: 'text/csv' });
        let url = window.URL.createObjectURL(blob);
        let element = document.createElement('a');
        element.setAttribute('href', url);
        element.setAttribute('download', "debugDB.csv");
        element.click();
      }
      const showDebug = () => {
        if(!debugStat) {
          debugInterval.value = 10;
          debugDuration.value = 10;
        }
        debugW.value = true;
      }
      const hideDebug = () => {
        debugW.value = false;
      }
      const sleep = (milliseconds) => {
        return new Promise(resolve => setTimeout(resolve, milliseconds));
      }
      const exeDebug = () => {
        if(currentQuery.debugData.length > 0) {
          if(debugStat.value) {
            debugStat.value = false;
            clearInterval(debugTimer);
            debugTimer = null;
            ElMessage("測試結束！")
          } else {
            ElMessage("測試開始！你可以關閉設定對話框，到主畫面查看送出狀態")
            if(debugInterval.value > 60) {
              debugInterval.value = 60;
            }
            if(debugInterval.value < 10) {
              debugInterval.value = 10;
            }
            if(sleepTime.value > debugInterval.value) {
              sleepTime.value = debugInterval.value;
            }
            if(sleepTime.value < 0) {
              sleepTime.value = 0;
            }
            if(debugDuration.value > 180) {
              debugDuration.value = 180;
            }
            if(debugDuration.value < 5) {
              debugDuration.value = 5;
            }
            debugStat.value = true;
            debugPrecentage.value = 0;
            clearInterval(debugTimer);
            let debugData = _.shuffle(currentQuery.debugData)
            let startTime = new Date().getTime();
            let exeTimes = 0;
            let exeTimer = debugDuration.value * 1000 * 60;
            let leftTimer = debugDuration.value * 1000 * 60;
            debugTimer = setInterval(() => {
              let now = new Date().getTime();
              if(!connecting.value) {
                if(leftTimer > 0) {
                  exeTimes++;
                  leftTimer = exeTimer - (now-startTime);
                  let timeLeft = leftTimer >= 0 ? leftTimer : 0;
                  debugPrecentage.value = Math.floor(((exeTimer - timeLeft) / exeTimer) * 100)
                  let testData = debugData[exeTimes % debugData.length]
                  for(let i=0;i<currentQuery.inputColumns.length; i++) {
                    let testValue = testData[currentQuery.inputColumns[i].name];
                    if(testValue !== undefined) {
                      let writeFlag = true;
                      if(formatDetector("S", currentQuery.inputColumns[i])) {
                        if(sufferSelect.value) {
                          writeFlag = false;
                        }
                      }
                      if(writeFlag) {
                        currentQuery.inputColumns[i].value = testValue;
                      }
                    } else {
                      if(sufferSelect.value) {
                        if(formatDetector("S", currentQuery.inputColumns[i])) {
                          let selections = _.shuffle(_.split(currentQuery.inputColumns[i].format, ";"));
                          if(selections.length > 0) {
                            currentQuery.inputColumns[i].value = selections[0];
                          }
                        }
                      }
                    }
                  }
                  sendInsert();
                  sleep(sleepTime.value);
                } else {
                  clearInterval(debugTimer);
                  debugTimer = null;
                  debugStat.value = false;
                  ElMessage("測試結束！")
                }
              } else {
                ElMessage("上一次通訊尚未結束，跳過這次定時器")
              }
            }, debugInterval.value * 1000)
          }
        } else {
          ElMessage("沒測試資料啊，你怎麼進來的？")
        }
      }
      const previewData = computed(() => {
        return currentQuery.debugData.length > 0 ? JSON.stringify(currentQuery.debugData[0]) : "無資料";
      });
      onMounted(() => {
        colors = randomColor({
          count: 30,
          luminosity: 'dark',
          format: 'rgb'
        });
        ElMessage("下載表單列表中");
        google.script.run.withSuccessHandler((list) => {
          resetMsg();
          sheetList.value = list;
          changeStep("表單列表", "process", "wait", "wait");
        })
        .withFailureHandler((data) => {
          resetMsg();
          sysMsg.type = "error";
          sysMsg.message = data;
          sysMsg.title = "主機端發生錯誤";
          sysMsg.status = true;
        })
        .buildList();
      });
      return {
        sheetList, changeStep, dialogs, stepIndicator, dateConverter, sysMsg, tagFinder, removeTag, currentQuery, checkAuth, connecting, reloadPage, checkLogin, openSheet, formatDetector, valField, sendInsert, checkInsert, viewableW, selectionMove, endSelection, startSelection, formatHelper, chooseSelection, downloadDebug, debugW, showDebug, hideDebug, exeDebug, debugDuration, debugInterval, debugStat, sleepTime, previewData, sufferSelect, debugPrecentage, progressColors
      };
    }
  }
</script>