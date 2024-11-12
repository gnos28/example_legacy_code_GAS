const API_KEY = ""

const today = new Date()
const todayUStime = today.getTime()
const todayUS = new Date(todayUStime + 6 * 60 * 60 * 1000) // risque de bug avec passage heure hiver / été

const currentDay = todayUS.getDate()
const currentMonth = todayUS.getMonth() + 1
const currentYear = todayUS.getFullYear()
const currentHour = todayUS.getHours()
const currentMinute = todayUS.getMinutes()
const currentWeekDay = getWeekDay(todayUS)
const currentWeek = iSO8601_week_no(todayUS)

const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()

function getWeekDay(date) {
  let weekDay = date.getDay()
  if (weekDay == 0)
    weekDay = 7

  return weekDay
}

function emailsAlert(previousData, newData) {

  scriptTime = new Date()
  console.log("emailAlert getSheetByName", scriptTime.getTime() - todayUStime)

  let alertSheet = activeSpreadsheet.getSheetByName("FOLLOW UP")

  scriptTime = new Date()
  console.log("emailAlert getRange", scriptTime.getTime() - todayUStime)

  let alertRange = alertSheet.getRange(1, 1, alertSheet.getMaxRows(), alertSheet.getMaxColumns())

  scriptTime = new Date()
  console.log("emailAlert getValues", scriptTime.getTime() - todayUStime)

  let alertValues = alertRange.getValues()

  scriptTime = new Date()
  console.log("emailAlert build alerts", scriptTime.getTime() - todayUStime)

  let alerts = []
  for (let i = 1; i < alertValues[0].length; i++)
    if (alertValues[0][i] !== "") {
      let email = alertValues[0][i]
      let threshold = alertValues[1][i]

      let assetList = []
      for (let j = 2; j < alertValues.length; j++)
        if (alertValues[j][i] !== "")
          assetList.push(alertValues[j][i])

      alerts.push({
        email: email,
        threshold: threshold,
        assets: assetList
      })
    }

  scriptTime = new Date()
  console.log("emailAlert build priceVariations", scriptTime.getTime() - todayUStime)

  alerts.forEach(alert => {
    scriptTime = new Date()
    console.log("emailAlert loop alert [start] ", alert.email, scriptTime.getTime() - todayUStime)
    let priceVariations = []
    alert.assets.forEach(asset => {
      let previousPrice = false
      let previousVolume = false

      for (let i = 0; i < previousData.length && !previousPrice; i++)
        if (previousData[i][2] === asset) {
          previousPrice = previousData[i][6]
          previousVolume = previousData[i][7]
        }

      let newPrice = false
      let newVolume = false

      for (let i = 0; i < newData.length && !newPrice; i++)
        if (newData[i][2] === asset) {
          newPrice = newData[i][6]
          newVolume = newData[i][7]
        }

      if (previousPrice !== false && newPrice !== false)
        priceVariations.push({
          asset: asset,
          priceEvo: newPrice / previousPrice - 1,
          volumeEvo: newVolume / previousVolume - 1,
        })
    })

    let emailBody = ""
    let emailTitle = ""

    priceVariations.sort((a, b) => {
      return b.priceEvo - a.priceEvo;
    });

    scriptTime = new Date()
    console.log("emailAlert build emailBody", scriptTime.getTime() - todayUStime)

    let fontColor = "green"
    priceVariations.forEach(price => {
      if (price.priceEvo >= alert.threshold || price.priceEvo <= -alert.threshold) {


        let priceEvo = Math.round((price.priceEvo * 100 + Number.EPSILON) * 100) / 100
        if (priceEvo > 0)
          priceEvo = "+" + priceEvo
        else {
          if (fontColor == "green")
            emailBody += "<br>"

          fontColor = "red"
        }

        let volumeEvo = Math.round((price.volumeEvo * 100 + Number.EPSILON) * 100) / 100
        if (volumeEvo > 0)
          volumeEvo = "+" + volumeEvo

        if (emailTitle !== "" && emailTitle.length < 120)
          emailTitle += " | "
        if (emailTitle.length < 120)
          emailTitle += `${price.asset}${priceEvo}%`
        else
          if(emailTitle.charAt(emailTitle.length-1) !== ".")
            emailTitle += " ..."
        emailBody += `<font color="${fontColor}"><b>${price.asset} : ${priceEvo}%</b> (vol : ${volumeEvo}%)</font><br>`
      }
    })
    scriptTime = new Date()
    console.log("emailAlert send email", scriptTime.getTime() - todayUStime)

    if (emailTitle !== "")
      MailApp.sendEmail(
        {
          to: alert.email,
          subject: emailTitle,
          htmlBody: emailBody,
        });

    scriptTime = new Date()
    console.log("emailAlert loop alert [end] ", scriptTime.getTime() - todayUStime)
  })
}

function betaTest() {
  let graphSheet = activeSpreadsheet.getSheetByName("GRAPH")
  let dataSheet = activeSpreadsheet.getSheetByName("HOURLY_ARCHIVE")

  let symbolListRange = graphSheet.getRange("E38:E47")
  let symbolListValues = symbolListRange.getValues()

  let indicesRange = graphSheet.getRange("E8:E10")
  let indicesValues = indicesRange.getValues()

  let dataSymbolRange = dataSheet.getRange("C:C")
  let dataSymbolValues = dataSymbolRange.getValues()

  let dataDateRange = dataSheet.getRange("F:F")
  let dataDateValues = dataDateRange.getValues()

  let dataPriceRange = dataSheet.getRange("H:H")
  let dataPriceValues = dataPriceRange.getValues()

  beta(symbolListValues, indicesValues, dataSymbolValues, dataDateValues, dataPriceValues)
}

export function beta_legacy(symbolList, indices, dataSymbol, dataDate, dataPrice) {
  let historySymbol = []
  let historyIndices = []
  let historySymbolUnsorted = []
  let historyIndicesUnsorted = []

  symbolList.forEach(symbol => {
    historySymbolUnsorted.push([])
  })

  indices.forEach(symbol => {
    historyIndicesUnsorted.push([])
  })

  dataSymbol.forEach((ticker, indexData) => {
    symbolList.forEach((symbol, indexSymbol) => {
      if (symbol[0] === ticker[0] && symbol[0] !== "") {
        historySymbolUnsorted[indexSymbol].push({
          date: dataDate[indexData][0].getTime(),
          price: dataPrice[indexData][0],
          change: ""
        })
      }
    })
    indices.forEach((indice, indexIndice) => {
      if (indice[0] === ticker[0] && indice[0] !== "") {
        historyIndicesUnsorted[indexIndice].push({
          date: dataDate[indexData][0].getTime(),
          price: dataPrice[indexData][0],
          change: ""
        })
      }
    })
  })

  historySymbolUnsorted.forEach((symbol, index) => {
    symbol.sort((a, b) => {
      return a.date - b.date;
    });
    historySymbol.push(symbol)
  })

  historyIndicesUnsorted.forEach((indice, index) => {
    indice.sort((a, b) => {
      return a.date - b.date;
    });
    historyIndices.push(indice)
  })


  let covariances = []
  let varianceIndicesArray = []

  historySymbol.forEach((symbol, indexSymbol) => {
    console.log("***** indexSymbol", indexSymbol)

    let globalIndicesChangeAverage = 0
    let nbIndicesChange = 0
    let indiceLength = 0

    let filteredSymbol = []
    let filteredHistoryIndices = []
    historyIndices.forEach((indice) => {
      filteredHistoryIndices.push([])
    })

    // ne garder que les dates communes
    historyIndices.forEach((indice, indexIndice) => {
      for (let i = 0; i < indice.length; i++) {
        let dateFound = false
        for (let j = 0; j < symbol.length && !dateFound; j++)
          if (symbol[j].date === indice[i].date)
            dateFound = true

        if (dateFound)
          filteredHistoryIndices[indexIndice].push(indice[i])
      }
    })

    for (let i = 0; i < symbol.length; i++) {
      let dateFound = false
      for (let k = 0; k < filteredHistoryIndices.length; k++)
        for (let j = 0; j < filteredHistoryIndices[k].length && !dateFound; j++)
          if (filteredHistoryIndices[k][j].date === symbol[i].date)
            dateFound = true

      if (dateFound)
        filteredSymbol.push(symbol[i])
    }

    // recalculer variance indice par rapport à ce symbol
    filteredHistoryIndices.forEach((indice, indexIndice) => {

      for (let i = 1; i < indice.length; i++) {

        let change = (indice[i].price / indice[i - 1].price) - 1
        if (!isNaN(change)) {
          filteredHistoryIndices[indexIndice][i].change = change
          globalIndicesChangeAverage += change
          nbIndicesChange++
        }
      }

      if (indice.length > indiceLength)
        indiceLength = indice.length
    })

    globalIndicesChangeAverage = globalIndicesChangeAverage / nbIndicesChange
    let varianceIndices = 0
    let nbVarianceFound = 0

    for (let i = 1; i < indiceLength; i++) {
      let sumChanges = 0
      let nbIndiceInput = 0
      for (let j = 0; j < filteredHistoryIndices.length; j++) {
        if (i < filteredHistoryIndices[j].length)
          if (!isNaN(filteredHistoryIndices[j][i].change)) {
            sumChanges += filteredHistoryIndices[j][i].change
            nbIndiceInput++
          }
      }
      let avgChange = sumChanges / nbIndiceInput
      if (!isNaN(avgChange)) {
        varianceIndices += Math.pow(avgChange + globalIndicesChangeAverage, 2)
        nbVarianceFound++
      }
    }

    varianceIndicesArray.push(varianceIndices / nbVarianceFound)
    console.log("varianceIndices", varianceIndices)
    console.log("nbVarianceFound", nbVarianceFound)
    console.log("varianceIndices / nbVarianceFound", varianceIndices / nbVarianceFound)

    let symbolChangeAverage = 0
    let nbSymbolChangeFound = 0

    // *** calcul covariance
    // calcul moyenne
    for (let i = 1; i < filteredSymbol.length; i++) {
      let change = (filteredSymbol[i].price / filteredSymbol[i - 1].price) - 1
      filteredSymbol[i].change = change
      if (!isNaN(change)) {
        symbolChangeAverage += change
        nbSymbolChangeFound++
      }
    }

    symbolChangeAverage = symbolChangeAverage / nbSymbolChangeFound

    let covariance = 0
    let covarianceDenominator = 0

    for (let i = 1; i < filteredSymbol.length; i++) {
      let indiceLoopAverage = 0
      let nbIndiceFound = 0
      filteredHistoryIndices.forEach(indice => {

        if (i < indice.length) {
          if (!isNaN(indice[i].change)) {
            indiceLoopAverage += indice[i].change
            nbIndiceFound++
          }
        }
      })

      console.log("+ + + + indiceLoopAverage", indiceLoopAverage, nbIndiceFound, indiceLoopAverage / nbIndiceFound)
      indiceLoopAverage = indiceLoopAverage / nbIndiceFound

      if (!isNaN(filteredSymbol[i].change) && !isNaN(symbolChangeAverage) && !isNaN(indiceLoopAverage) && !isNaN(globalIndicesChangeAverage)) {
        covariance += ((filteredSymbol[i].change + symbolChangeAverage) * (indiceLoopAverage + globalIndicesChangeAverage))
        console.log("* * * * covariance", covariance, "|", filteredSymbol[i].change, "|", symbolChangeAverage, "|", indiceLoopAverage, "|", globalIndicesChangeAverage)
        covarianceDenominator++
      }
    }
    console.log("symbolChangeAverage", symbolChangeAverage)
    console.log("globalIndicesChangeAverage", globalIndicesChangeAverage)

    covariances.push(covariance / covarianceDenominator)
    console.log("covariance", covariance)
    console.log("covarianceDenominator", covarianceDenominator)

  })

  let outputArray = []

  covariances.forEach((covarianceB, index) => {
    if (!isNaN(covarianceB))
      outputArray.push([covarianceB / varianceIndicesArray[index]])
    else
      outputArray.push([""])
  })

  return outputArray
}

export function RSI_legacy(symbolList, dataSymbol, dataDate, dataPrice) {
  let priceList = []
  let upList = []
  let downList = []

  symbolList.forEach(symbol => {
    priceList.push([])
  })

  for (let i = 0; i < dataSymbol.length; i++) {

    symbolList.forEach((symbol, index) => {

      if (dataSymbol[i][0] === symbol[0] && dataSymbol[i][0] !== "" && dataDate[i][0] !== "") {
        priceList[index].push({
          date: dataDate[i][0].getTime(),
          price: dataPrice[i][0],
          rsi: "",
        })
      }
    })
  }

  let returnArray = []

  //return JSON.stringify(priceList[0])

  priceList.forEach(price => {

    if (price.length >= 14) {
      price.sort((a, b) => {
        return a.date - b.date;
      });

      let rsiPeriod = price.length - 1
      let rsiPeriodD = 2 / (rsiPeriod + 1)

      for (let i = price.length - 14; i < price.length; i++) {
        let change = (price[i].price / price[i - 1].price) - 1
        if (change > 0)
          upList.push(change)
        else
          downList.push(change)
      }

      let upListAvg = 0
      for (let i = 0; i < upList.length; i++) {
        upListAvg = rsiPeriodD * upList[i] + (1 - rsiPeriodD) * upListAvg
      }

      let downListAvg = 0
      for (let i = 0; i < downList.length; i++) {
        downListAvg = rsiPeriodD * downList[i] + (1 - rsiPeriodD) * downListAvg
      }

      let rs = upListAvg / Math.abs(downListAvg)

      price.rsi = 100 - (100 / (1 + rs))

      returnArray.push([price.rsi])
    }
    else
      returnArray.push([''])
  })

  return returnArray

}

function iSO8601_week_no(dt) {
  var tdt = new Date(dt.valueOf());
  var dayn = (dt.getDay() + 6) % 7;
  tdt.setDate(tdt.getDate() - dayn + 3);
  var firstThursday = tdt.valueOf();
  tdt.setMonth(0, 1);
  if (tdt.getDay() !== 4) {
    tdt.setMonth(0, 1 + ((4 - tdt.getDay()) + 7) % 7);
  }
  return 1 + Math.ceil((firstThursday - tdt) / 604800000);
}

function getListingLatest() {
  let header = { // == headers (pour google) ou header
    'Content-Type': 'application/json; charset=UTF-8',
    'Accept': 'application/json; charset=UTF-8',
    'X-CMC_PRO_API_KEY': API_KEY,
    "_method": "GET",
    // 'start':1,
    // 'limit':200,
  }



  let headerString = JSON.stringify(header)


  const url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest?start=1&limit=200'

  let params = {
    'method': 'GET',
    'contentType': 'application/json; charset=utf-8',
    'headers': header, // google
    'muteHttpExceptions': true
  }

  let fetch = UrlFetchApp.fetch(url, params)

  return fetch
}

function getPreviousData(sheetName) {
  let storeSheet = activeSpreadsheet.getSheetByName(sheetName)
  let storeRange = storeSheet.getRange(2, 1, storeSheet.getMaxRows() - 1, storeSheet.getMaxColumns())
  let storeValues = storeRange.getValues()

  let outputArray = []
  // remove empty lines
  storeValues.forEach(line => {
    if (line[2] !== "")
      outputArray.push(line)
  })

  return outputArray
}

function debugStoreData() {
  let previousData = getPreviousData("OUTPUT API")

  storeData("DAILY_ARCHIVE_TEST", previousData, 144 * 24 * 60 * 60 * 1000, "day")

}

function storeData(sheetName, outputArray, keepTime, override) {

  scriptTime = new Date()
  console.log("storeData getRange", scriptTime.getTime() - todayUStime)

  let storeSheet = activeSpreadsheet.getSheetByName(sheetName)

  let storeRange
  if (!override)
    storeRange = storeSheet.getRange(1, 1, storeSheet.getMaxRows(), 6)
  else
    storeRange = storeSheet.getRange(1, 1, storeSheet.getMaxRows() - 300, 6)

  let storeValues = storeRange.getValues()

  scriptTime = new Date()
  console.log("storeData searchEmptyRanges", scriptTime.getTime() - todayUStime)

  // rechercher des blocs de lignes vides dans storeValues
  let startEmptyIndex = false
  let currentEmptyIndex = false
  let emptyRanges = []
  storeValues.forEach((line, index) => {
    if (line[2] === "") {
      if (!startEmptyIndex) {
        startEmptyIndex = index
        currentEmptyIndex = index
      }
      else
        currentEmptyIndex = index
    }
    else {
      if (startEmptyIndex !== false) {
        emptyRanges.push({
          start: startEmptyIndex,
          end: currentEmptyIndex
        })
        startEmptyIndex = false
        currentEmptyIndex = false
      }
    }
  })

  if (storeValues[storeValues.length - 1][2] === "")
    emptyRanges.push({
      start: startEmptyIndex,
      end: currentEmptyIndex
    })

  scriptTime = new Date()
  console.log("storeData add newData", scriptTime.getTime() - todayUStime)

  // ajouter nouvelles valeurs sur lignes vides
  let outputArrayIndex = 0
  let outputArraySize = outputArray.length
  emptyRanges.forEach(range => {
    if (outputArrayIndex < outputArraySize) {

      console.log("**** outputArrayIndex < outputArraySize", outputArrayIndex, outputArraySize)

      let rangeSize = range.end - range.start + 1

      console.log("rangeSize", rangeSize)

      let tinyStoreRange = storeSheet.getRange(range.start + 1, 1, Math.min(outputArraySize, rangeSize), outputArray[0].length)
      let tinyStoreValues = tinyStoreRange.getValues()

      console.log("tinyStoreRange", range.start + 1, 1, Math.min(outputArraySize, rangeSize), outputArray[0].length)

      let nbValuesStored = 0
      for (let i = 0; i < tinyStoreValues.length && i + outputArrayIndex < outputArraySize; i++) {
        for (let j = 0; j < outputArray[0].length; j++) {
          tinyStoreValues[i][j] = outputArray[outputArrayIndex + i][j]
        }
        nbValuesStored++
      }

      console.log("tinyStoreValues.length", tinyStoreValues.length)

      tinyStoreRange.setValues(tinyStoreValues)
      outputArrayIndex += nbValuesStored

      console.log("outputArrayIndex", outputArrayIndex)

      console.log("**** end loop")
    }
  })

  scriptTime = new Date()
  console.log("storeData remove old Data", scriptTime.getTime() - todayUStime)

  // virer les valeurs > keepTime
  let todayTime = today.getTime()
  let startDelIndex = false
  let currentDelIndex = false
  storeValues.forEach((line, index) => {
    if (index > 0 && line[2] !== "" && line[5] !== "") {
      let dateImport = line[5].getTime()

      let keepTimeCheck = todayTime - dateImport > keepTime
      let overrideCheck = false
      // switch (override) {
      //   case "hour":
      //     if (line[10] == currentYear && line[11] == currentMonth && line[12] == currentDay && line[13] == currentHour)
      //       overrideCheck = true
      //     break;
      //   case "day":
      //     if (line[10] == currentYear && line[11] == currentMonth && line[12] == currentDay)
      //       overrideCheck = true
      //     break;
      //   case "week":
      //     if (line[10] == currentYear && line[16] == currentWeek)
      //       overrideCheck = true
      //     break;
      //   case "month":
      //     if (line[10] == currentYear && line[11] == currentMonth)
      //       overrideCheck = true
      //     break;
      // }

      if (keepTimeCheck || overrideCheck) {
        if (!startDelIndex) {
          startDelIndex = index
          currentDelIndex = index
        }

        if (index !== currentDelIndex + 1) // delete range a init a new range
        {
          let clearRange = storeSheet.getRange(startDelIndex + 1, 1, (currentDelIndex - startDelIndex) + 1, storeSheet.getMaxColumns())
          clearRange.clearContent()
          startDelIndex = false
          currentDelIndex = false
        }
        else {
          currentDelIndex = index
        }
      }
    }
  })

  // reconstruire les formules
  // let formulas = [
  //   // '=if(C2<>"";year($F2);"")',
  //   // '=if(D2<>"";MONTH($F2);"")',
  //   // '=if(E2<>"";DAY($F2);"")',
  //   // '=if(F2<>"";HOUR($F2);"")',
  //   // '=if(G2<>"";MINUTE($F2);"")',
  //   '=if($C2<>"";weekday($F2;2);"")',
  //   '=if($C2<>"";ISOWEEKNUM($F2);"")'
  // ];

  // for (let i = 0; i < formulas.length; i++) {
  //   let formulaRange = storeSheet.getRange(2, 22 + i, storeSheet.getMaxRows() - 1, 1)
  //   formulaRange.setFormula(formulas[i])
  // }

  scriptTime = new Date()
  console.log("storeData finish", scriptTime.getTime() - todayUStime)
}

export function hourlyRefresh() {

  console.log("start")
  let graphSheet = activeSpreadsheet.getSheetByName("GRAPH")
  let graphRange = graphSheet.getRange(1, 1, 1, 1)
  graphRange.setValues([["REFRESHING DATAS ..."]])

  let response = getListingLatest()
  let responseCode = response.getResponseCode()

  if (responseCode == 200) {

    // récupérer les indices hors crypto (SP, USDindex, ...)
    let otherSheet = activeSpreadsheet.getSheetByName("OTHERS")

    let otherRange = otherSheet.getRange(1, 1, otherSheet.getMaxRows(), 5)
    let otherValues = otherRange.getValues()

    // variables pour top10 & top200
    let top10 = {
      name: 'top10 crypto',
      symbol: '_TOP10',
      price: 0,
      volume_24h: 0,
      market_cap: 0,
      market_cap_dominance: 0,
    }

    let top200 = {
      name: 'top200 crypto',
      symbol: '_TOP200',
      price: 0,
      volume_24h: 0,
      market_cap: 0,
      market_cap_dominance: 0,
    }

    let json = JSON.parse(response.getContentText())

    let outputArray = []

    json.data.forEach((crypto, index) => {
      let currency = Object.keys(crypto.quote)[0]

      // top 10
      if (index < 10) {
        top10.price = (top10.price * index + crypto.quote[currency].price) / (index + 1) // price average
        top10.volume_24h = top10.volume_24h + crypto.quote[currency].volume_24h
        top10.market_cap = top10.market_cap + crypto.quote[currency].market_cap
        top10.market_cap_dominance = top10.market_cap_dominance + crypto.quote[currency].market_cap_dominance
      }

      // top 200
      top200.price = (top200.price * index + crypto.quote[currency].price) / (index + 1) // price average
      top200.volume_24h = top200.volume_24h + crypto.quote[currency].volume_24h
      top200.market_cap = top200.market_cap + crypto.quote[currency].market_cap
      top200.market_cap_dominance = top200.market_cap_dominance + crypto.quote[currency].market_cap_dominance

      // retrait des chiffres abberants
      let volume_24h = crypto.quote[currency].volume_24h
      if (volume_24h > 1000000000000000)
        volume_24h = ""

      let volume_change_24h = crypto.quote[currency].volume_change_24h
      if (volume_change_24h > 1000000)
        volume_change_24h = ""

      outputArray.push([
        crypto.id,
        crypto.name,
        crypto.symbol,
        crypto.circulating_supply,
        "",
        today,
        crypto.quote[currency].price,
        volume_24h,
        crypto.quote[currency].market_cap,
        crypto.quote[currency].market_cap_dominance,
        currentYear,
        currentMonth,
        currentDay,
        currentHour,
        currentMinute,
        currentWeekDay,
        currentWeek
      ])
    })

    // top 10
    outputArray.push([
      "",
      top10.name,
      top10.symbol,
      "",
      "",
      today,
      top10.price,
      top10.volume_24h,
      top10.market_cap,
      top10.market_cap_dominance,
      currentYear,
      currentMonth,
      currentDay,
      currentHour,
      currentMinute,
      currentWeekDay,
      currentWeek
    ])

    // top 200
    outputArray.push([
      "",
      top200.name,
      top200.symbol,
      "",
      "",
      today,
      top200.price,
      top200.volume_24h,
      top200.market_cap,
      top200.market_cap_dominance,
      currentYear,
      currentMonth,
      currentDay,
      currentHour,
      currentMinute,
      currentWeekDay,
      currentWeek
    ])

    otherValues.forEach(line => {
      if (line[0] !== "") {
        let last_updated = line[3]
        if (line[4] != "OPEN")
          last_updated = "CLOSED"

        outputArray.push([
          "",
          line[1],
          line[0],
          "",
          last_updated,
          today,
          line[2],
          "",
          "",
          "",
          currentYear,
          currentMonth,
          currentDay,
          currentHour,
          currentMinute,
          currentWeekDay,
          currentWeek
        ])
      }
    })

    let scriptTime = new Date()
    console.log("previousData", scriptTime.getTime() - todayUStime)

    let previousData = getPreviousData("OUTPUT API")

    let previousHour = previousData[0][13]
    let previousDay = previousData[0][12]
    let previousWeek = previousData[0][16]
    let previousMonth = previousData[0][11]

    scriptTime = new Date()
    console.log("*** storeData OUTPUT API", scriptTime.getTime() - todayUStime)
    storeData("OUTPUT API", outputArray, 1000, false)

    scriptTime = new Date()
    console.log("*** storeData MINUTE ARCHIVE", scriptTime.getTime() - todayUStime)
    storeData("MINUTE_ARCHIVE", outputArray, 48 * 60 * 60 * 1000, false)

    scriptTime = new Date()
    console.log("*** storeData HOURLY ARCHIVE", scriptTime.getTime() - todayUStime)
    if (previousHour != currentHour)
      storeData("HOURLY_ARCHIVE", previousData, 6.5 * 24 * 60 * 60 * 1000, "hour")

    scriptTime = new Date()
    console.log("*** storeData DAILY_ARCHIVE", scriptTime.getTime() - todayUStime)
    if (previousDay != currentDay)
      storeData("DAILY_ARCHIVE", previousData, 144 * 24 * 60 * 60 * 1000, "day")

    scriptTime = new Date()
    console.log("*** storeData WEEKLY_ARCHIVE", scriptTime.getTime() - todayUStime)
    if (previousWeek != currentWeek)
      storeData("WEEKLY_ARCHIVE", previousData, 144 * 7 * 24 * 60 * 60 * 1000, "week")

    scriptTime = new Date()
    console.log("*** storeData MONTHLY_ARCHIVE", scriptTime.getTime() - todayUStime)
    if (previousMonth != currentMonth)
      storeData("MONTHLY_ARCHIVE", previousData, 12 * 365 * 24 * 60 * 60 * 1000, "month")

    scriptTime = new Date()
    console.log("*** email alert", scriptTime.getTime() - todayUStime)
    emailsAlert(previousData, outputArray)

    scriptTime = new Date()
    console.log("finish !", scriptTime.getTime() - todayUStime)


  }

  graphRange.clearContent()
  SpreadsheetApp.flush()
}
