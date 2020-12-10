## This project includes two optimization models for portfolio construction

### Generic Convex Optimization

CVXPY is chosen after comparing  it against CVXOPT for better stability and flexibility.

Main part of the model is to formulate the target into an objective function and constraints expressed as equations. For example

    # obj
    objective = cp.norm(x - target_mkv) # minimize distance to target portfolio
    objective = cp.norm((x - cur_hold_mkv)) # minimize turnover
    objective = cp.sum(order_mkv * total_cost) # minimize trading cost
    # constrains
    cons_basic = [cp.abs(x) <= max_dollar_position,
                  cp.sum(x * l) == cap_l,
                  (x * l) >= 1e-9, #1e-9
                  (x * s) <= -1e-9, #1e-9
                  cp.sum(x * (beta * ls)) <= exposure_beta_range[1],
                  cp.sum(x * (beta * ls)) >= exposure_beta_range[0],
                  cp.sum(x * ls) <= exposure_dollar_range[1],
                  cp.sum(x * ls) >= exposure_dollar_range[0],
                  cp.abs(x * not_ls) <= 1e-9,
                  ]  
    
    cons_liquid = [cp.abs(x - cur_hold_mkv) <= liquid_mkv] # liquidity constraint
    cons_subind = get_subind_constr_abs (df_daily_stock_variable, 
                                         x, ls, beta, cap_l, subind_range) # exposure at different sub-group levels
The advantage of the model is it can take hundreds of parameters into account with reasonable efficiency and objective can be highly customized. The drawback is complex constraints can make the optimal surface discrete and solution can not be found. In practice, number of parameters in the optimization objective and constraints should be kept as minimal as possible. 



### Critical Line Algorithm 

Simple and Robust model to find solution for mean-variance optimization based on construction of efficient frontier. It is useful for simple portfolio optimization such as maximizing return and Sharpe, minimizing risk. Drawback is it cannot handle complex target and constraints. 



