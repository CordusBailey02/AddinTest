// ######################## BOND PAYMENT CONSTANTS ########################

// Bond Calculations
//const EMPLOYEE_COMMISSION_RATE = 0.02; // PLACE HOLDER <- converted to system variable later to be read from excel sheet
const COMMISSION_CAP_CONSTANT = 20000; // Used to calculate commission cap with commission rate

const PAYMENT_PLAN_THRESHOLD = 5000; // Bonds >= this go to payment plans, else commission
const COMMISSION_PLAN = "commision_plan";
const PAYMENT_PLAN = "payment_plan";

const MIN_DOWN_PAYMENT = 25; // Minimum commission payment

const FIRST_POWER_COST = 25; // Cost of first power -> used with formula
const ADDITIONAL_POWER_COST = 5; // Cost of each additional power -> used with formula

// ######################## BOND PAYMENT SHEET CALCULATION FUNCTIONS ########################

function calcPaymentPlanBond(ttl_bond, ttl_expense, travel, amt_collected, balance_owed, num_powers, commission_rate) {
  // Variables to return
  let percent_paid_on_bond = 0;
  let reach_payment_goal = 0;
  let employee_owed_balance = 0;
  let down_payment = 0;


  // Company cost
  const company_cost = (ttl_bond * 0.027) + ttl_expense + travel;
  
  // Net position
  const net_position = amt_collected - company_cost;

  // If net_position is NEGATIVE
  if(net_position < 0) {
    reach_payment_goal = balance_owed - Math.abs(net_position);

    down_payment = calcPowerCost(num_powers);

    employee_owed_balance = calcCappedCommission(ttl_bond, commission_rate) - down_payment

    percent_paid_on_bond = employee_owed_balance / reach_payment_goal;

  }
  else if(net_position >= 0) {
    down_payment = net_position * 0.25

    // If down payment is less than 25, make down payment 25
    if(down_payment < MIN_DOWN_PAYMENT) {
      down_payment = 25
    }

    employee_owed_balance = calcCappedCommission(ttl_bond, commission_rate) - down_payment;

    percent_paid_on_bond = employee_owed_balance / balance_owed;

  }

  // Conidition if bond >= 50k and % is < 20%, make their % 20%
  if(ttl_bond >= 50000 && percent_paid_on_bond < 0.20) {
      percent_paid_on_bond = 0.20
  }

  return {
    percent_paid_on_bond,
    reach_payment_goal,
    employee_owed_balance,
    down_payment,
  }

}

// ######################## BOND PAYMENT SHEET CALCULATION HELPER FUNCTIONS ########################

// Function to return the commision cap for specified rate
function getCommissionCap(commission_rate) {
  return COMMISSION_CAP_CONSTANT * commission_rate;
}

// Function to determine if commision cap or calculated value is used
function calcCappedCommission(ttl_bond, commission_rate) {
  const commission_cap = getCommissionCap(commission_rate);
  const raw = ttl_bond * commission_rate;
  return Math.min(commission_cap, raw);
}

function calcFullPayCommissionLT5(ttl_bond, commission_rate) {
  return Math.max(MIN_DOWN_PAYMENT, (ttl_bond * commission_rate))
}

function calcFullPayCommissionGTE5(ttl_bond, commission_rate) {
  return calcCappedCommission(ttl_bond, commission_rate);
}

// Function to calculate the cost for powers used
function calcPowerCost(powers) {
  return FIRST_POWER_COST + (ADDITIONAL_POWER_COST * (powers - 1))
}

function determineCommisionOrPaymentPlan(ttl_bond, balance_owed) {
  // If total bond is below payment plan threshold, then it goes to commision
  if(ttl_bond < PAYMENT_PLAN_THRESHOLD) {
    return COMMISSION_PLAN;
  }

  // If total bond is above payment plan threshold AND fully paid odd, then it goes to commision
  else if(ttl_bond > PAYMENT_PLAN_THRESHOLD && balance_owed == 0) {
    return COMMISSION_PLAN;
  }

  // If total bond is above payment plan threshold AND NOT fully paid off, then it goes to payment plan (PAYMENTS sheet)
  else if(ttl_bond > PAYMENT_PLAN_THRESHOLD && balance_owed != 0) {
    return PAYMENT_PLAN;
  }

  // Error case
  else {
    console.log("Error determining commission plan or payment plan...")
    console.log(`Total Bond Amount: ${ttl_bond}; Balance Owed: ${balance_owed}`)
  }
}


export {
  PAYMENT_PLAN_THRESHOLD,
  COMMISSION_PLAN,
  PAYMENT_PLAN,
  calcPaymentPlanBond,
  calcFullPayCommissionLT5,
  calcFullPayCommissionGTE5,
  determineCommisionOrPaymentPlan
};