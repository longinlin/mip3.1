## 一個生產程式的語言MIP                                                        <br>
  這裏有一個資料夾，裡面有一個MIP語言解譯器(MIP=Macro Interactive Processing)， <br>
  可以用來操控本地電腦或外地資料庫，輸出畫面結果。也可以是程式產生器，          <br>
  用來產出一整個資料夾的程式原始碼。                                            <br>
  .                                                                             <br>
  簡單用它的時候，你可以對它輸入一句接一句的指令。                              <br>
  擴大用它的時候，你可以巨集裡再叫用巨集(macro、宏、巨集)。                     <br>
  複雜用它的時候，你可以定義變數、迴圈、副程式。                                <br>
  .                                                                             <br>
  它是一個script解譯器，在前端顯示互動式網頁，也在後端連接資料庫；              <br>
  它是一個macro組合器，組合變數、一維文字序列(向量)、二維文字序列(矩陣)、       <br>
  文字區塊、含變數的模板 甚至 runTime才變化的模板。                             <br>
  .                                                                             <br>
  資料夾位置在 https://github.com/longinlin/MIP3.1 ， 複製到本機後，                <br>
  設定資料夾成為微軟IIS的虛擬目錄就可執行。免安裝。                             <br>
  以更少的寫字，表達更長的語意，提升資訊人員的生產力。                          <br>
  .                                                                             <br>
## 適用場合                                                                     <br>
 (1)生產大約類似的程式                                                          <br>
   多年以來，常見資訊人員日常就是下SQL指令，應付使用者的各種需求。              <br>
   使用者的需求變來變去很難寫成程式，如果能用簡語表達網頁輸入框，               <br>
   又能用簡語連接資料庫，這樣資訊人員才能正常上下班，這就是MIP的出發點。        <br>
   商用程式不僅是僵硬的增刪改查，實務上還有各種訊號聯動及資料轉移，             <br>
   所以MIP不是產生僵硬的框架，而是全面的減少編碼寫字，簡潔描述動作。            <br>
   本文將舉例兩個MIP寫的應用系統: 庫存撿貨系統，問卷評分系統。                  <br>
 (2)以組合代替寫作(或是代替複製貼上)                                            <br>
   物件導向流行以來，資訊界的程式語言越來越繁瑣，於是近年來又回頭流行巨集、     <br>
   模板、宏編程、元編程。MIP就是元編程語言，可以用來組合程式原始碼。            <br>
   你可以準備好模組文字，叫起MIP，由它產生多份文字檔。                          <br>
   也可以準備好模組文字，叫起MIP，由它不斷輸出新變數，變數又即時變更文字檔。    <br>
   .                                                                            <br>
   我們可以把MIP和java 混雜寫簡碼，輸出一個好幾倍長的java。                     <br>
   我們可以把MIP和C方言混雜寫簡碼，輸出一個好幾倍長的C程式。                    <br>
 (3)收攏多條向量，成為一個矩陣                                                  <br>
   物件導向的程式語越來越長了，程式中不斷宣告物件的屬性。                       <br>
   1個屬性佔1行，1個物件有9個屬性，8個物件就72行，這些物件都長得很像，          <br>
   讀完72行，程式重點還沒開始。何不把它們排成9乘8的矩陣，整齊清楚。             <br>
   用在商業程式，一行一行寫程式的一維思考，就可以轉為矩陣化的二維思考。         <br>
 (4)微縮程式碼，從真空管微縮到積體電路                                          <br>
   MIP不僅縮短一支程式，還可以縮短幾十支程式。當你有幾十支程式要寫，            <br>
   你找出這幾十支程式共同的部份，分離相異的部份，呼叫MIP就組合出幾十支程式。    <br>
