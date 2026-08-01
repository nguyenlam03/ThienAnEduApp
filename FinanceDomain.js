/**
 * Finance Domain
 * Các phép tính tài chính thuần, không đọc/ghi Google Sheet và không phụ thuộc giao diện.
 */
var FinanceDomain = (function () {
  function number(value) {
    var parsed = Number(value);
    return isFinite(parsed) ? parsed : 0;
  }

  function clamp(value, min, max) {
    return Math.max(min, Math.min(max, number(value)));
  }

  function calculatePricing(input) {
    input = input || {};
    var studentCount = Math.max(0, Math.floor(number(input.studentCount)));
    var currentAverageFee = Math.max(0, number(input.currentAverageFee));
    var plannedExpense = Math.max(0, number(input.plannedExpense));
    var variableExpense = Math.max(0, Math.min(plannedExpense, number(input.variableExpense)));
    var targetMargin = clamp(input.targetMargin, 0, 80);
    var collectionRate = clamp(input.collectionRate === undefined ? 100 : input.collectionRate, 1, 100);
    var denominator = studentCount * (collectionRate / 100) * (1 - targetMargin / 100);
    var requiredAverageFee = denominator > 0 ? plannedExpense / denominator : 0;
    var variablePerStudent = studentCount > 0 ? variableExpense / studentCount : 0;
    var fixedExpense = Math.max(0, plannedExpense - variableExpense);
    var contributionPerStudent = Math.max(0, currentAverageFee * collectionRate / 100 - variablePerStudent);
    var breakEvenStudents = contributionPerStudent > 0 ? Math.ceil(fixedExpense / contributionPerStudent) : 0;
    var scenarios = [0, 5, 8, 10].map(function (rate) {
      var averageFee = currentAverageFee * (1 + rate / 100);
      var revenue = averageFee * studentCount * collectionRate / 100;
      var profit = revenue - plannedExpense;
      return {
        rate: rate,
        averageFee: averageFee,
        revenue: revenue,
        profit: profit,
        margin: revenue > 0 ? profit * 100 / revenue : 0
      };
    });
    return {
      studentCount: studentCount,
      currentAverageFee: currentAverageFee,
      requiredAverageFee: requiredAverageFee,
      targetMargin: targetMargin,
      collectionRate: collectionRate,
      variablePerStudent: variablePerStudent,
      fixedTotal: fixedExpense,
      breakEvenStudents: breakEvenStudents,
      scenarios: scenarios
    };
  }

  function calculatePerformance(input) {
    input = input || {};
    var cashIncome = Math.max(0, number(input.cashIncome));
    var cashExpense = Math.max(0, number(input.cashExpense));
    var currentCash = number(input.currentCash);
    var revenueForecast = Math.max(0, number(input.revenueForecast));
    var accruedRevenue = Math.max(0, number(input.accruedRevenue));
    var plannedExpense = Math.max(0, number(input.plannedExpense));
    var remainingObligations = Math.max(0, number(input.remainingObligations));
    var reserveTarget = Math.max(0, number(input.reserveTarget));
    var ownerDraw = Math.max(0, number(input.ownerDraw));
    var projectedProfit = revenueForecast - plannedExpense;
    return {
      cashIncome: cashIncome,
      cashExpense: cashExpense,
      cashResult: cashIncome - cashExpense,
      currentCash: currentCash,
      revenueForecast: revenueForecast,
      accruedRevenueEstimate: accruedRevenue,
      plannedExpense: plannedExpense,
      remainingObligations: remainingObligations,
      estimatedProfit: accruedRevenue - plannedExpense,
      projectedProfit: projectedProfit,
      projectedMargin: revenueForecast > 0 ? projectedProfit * 100 / revenueForecast : 0,
      ownerDraw: ownerDraw,
      safeCash: currentCash - remainingObligations - reserveTarget
    };
  }

  return {
    calculatePricing: calculatePricing,
    calculatePerformance: calculatePerformance
  };
})();

function calculateFinancePricing(token, input) {
  SecurityService.requireSession(token, 'finance.read');
  return jsonResponse_(FinanceDomain.calculatePricing(input));
}
