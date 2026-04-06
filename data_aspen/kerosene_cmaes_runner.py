import csv
import os
import re
import time
from pathlib import Path

import numpy as np
import pythoncom
import win32com.client as win32
from pythoncom import com_error


DIALOG_TITLE_HINTS = (
    "프린터 연결을 기다리는 중",
    "waiting for printer connection",
)


class AspenSimulation:
    def __init__(self, AspenFileName, WorkingDirectoryPath, VISIBILITY=False, suppress_dialogs=False):
        pythoncom.CoInitialize()

        self.working_directory = Path(WorkingDirectoryPath).resolve()
        self.abs_path = self._resolve_aspen_path(AspenFileName, self.working_directory)
        self.last_run_error = None
        self.last_per_error = None
        self.dialog_suppression_enabled = bool(suppress_dialogs)

        print(f"Aspen working directory: {self.working_directory}")
        print(f"Aspen file: {self.abs_path}")

        os.chdir(self.working_directory)
        self.AspenSimulation = win32.DispatchEx("Apwn.Document")
        self.AspenSimulation.InitFromFile(str(self.abs_path))
        self.AspenSimulation.Visible = bool(VISIBILITY)
        self.DialogSuppression(self.dialog_suppression_enabled)
        print(f"Aspen loaded successfully. Visible={getattr(self.AspenSimulation, 'Visible', 'unknown')}")

    @staticmethod
    def _resolve_aspen_path(aspen_file_name, working_directory):
        candidate = Path(aspen_file_name)
        if not candidate.is_absolute():
            candidate = working_directory / candidate
        candidate = candidate.resolve()
        if not candidate.exists():
            raise FileNotFoundError(f"Aspen file not found: {candidate}")
        return candidate

    def CloseAspen(self):
        try:
            self.AspenSimulation.Close(False)
        except Exception:
            try:
                self.AspenSimulation.Quit()
            except Exception:
                pass
        print("Aspen should be closed now")

    def DialogSuppression(self, enabled=True):
        self.dialog_suppression_enabled = bool(enabled)
        if hasattr(self.AspenSimulation, "SuppressDialogs"):
            self.AspenSimulation.SuppressDialogs = bool(enabled)
        elif hasattr(self.AspenSimulation, "UIDisableDialogs"):
            self.AspenSimulation.UIDisableDialogs = bool(enabled)

    def DismissAspenDialogs(self):
        try:
            import win32con
            import win32gui
        except Exception:
            return []

        closed_titles = []

        def enum_handler(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                title = win32gui.GetWindowText(hwnd) or ""
                if not title:
                    return
                title_lower = title.lower()
                if any(hint.lower() in title_lower for hint in DIALOG_TITLE_HINTS):
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    closed_titles.append(title)
            except Exception:
                return

        try:
            win32gui.EnumWindows(enum_handler, None)
        except Exception:
            return []

        if closed_titles:
            unique_titles = sorted(set(closed_titles))
            print(f"Auto-closed Aspen dialog(s): {unique_titles}")
            return unique_titles
        return []

    @property
    def BLK(self):
        return self.AspenSimulation.Tree.Elements("Data").Elements("Blocks")

    @property
    def STRM(self):
        return self.AspenSimulation.Tree.Elements("Data").Elements("Streams")

    def read_block_input_value(self, block_name, input_name, fallback=None):
        try:
            block = self.BLK.Elements(block_name)
            if block is None:
                return fallback
            input_section = block.Elements("Input")
            if input_section is None:
                return fallback
            node = input_section.Elements(input_name)
            if node is None:
                return fallback
            value = getattr(node, "Value", None)
            if value is None:
                return fallback
            return value
        except Exception:
            return fallback

    def get_kerosene_purity(self, stream_name, start_c=8, end_c=16):
        try:
            tree = self.AspenSimulation.Tree
            total_mass_node = tree.FindNode(rf"\Data\Streams\{stream_name}\Output\RES_MASSFLOW")
            total_mass_flow = float(total_mass_node.Value) if total_mass_node and total_mass_node.Value is not None else 0.0
            if total_mass_flow <= 0.0:
                return 0.0

            target_mass = 0.0
            for carbon_number in range(start_c, end_c + 1):
                node = tree.FindNode(rf"\Data\Streams\{stream_name}\Output\MASSFLOW3\C{carbon_number}")
                if node is not None and node.Value is not None:
                    target_mass += float(node.Value)
            return target_mass / total_mass_flow
        except Exception as err:
            print(f"Error calculating purity: {err}")
            return 0.0

    def get_ft_selectivity(self, target_stream, product_streams, start_c=8, end_c=16):
        try:
            tree = self.AspenSimulation.Tree
            total_product_mass = 0.0
            for stream_name in product_streams:
                node = tree.FindNode(rf"\Data\Streams\{stream_name}\Output\RES_MASSFLOW")
                if node is not None and node.Value is not None:
                    total_product_mass += float(node.Value)
            if total_product_mass <= 0.0:
                return 0.0

            target_mass = 0.0
            for carbon_number in range(start_c, end_c + 1):
                node = tree.FindNode(rf"\Data\Streams\{target_stream}\Output\MASSFLOW3\C{carbon_number}")
                if node is not None and node.Value is not None:
                    target_mass += float(node.Value)
            return target_mass / total_product_mass
        except Exception as err:
            print(f"Error calculating selectivity: {err}")
            return 0.0

    def get_utility_cost(self):
        try:
            tree = self.AspenSimulation.Tree
            node = tree.FindNode(r"\Data\Results Summary\Utility-Sum\Output\TOTUTCOST")
            if node is not None and node.Value is not None:
                return float(node.Value)
            return 0.0
        except Exception as err:
            print(f"Error extracting utility cost: {err}")
            return 0.0

    def BLK_RPLUG_Set_T_SPEC_Constant_Temp(self, Blockname, ReactorTemperature):
        block_input = self.BLK.Elements(Blockname).Elements("Input")
        try:
            block_input.Elements("OPT_TSPEC").Value = "CONST-TEMP"
        except Exception:
            pass

        for node_name in ("CTEMP", "REAC_TEMP", "TEMP"):
            try:
                node = block_input.Elements(node_name)
                if node is not None:
                    node.Value = float(ReactorTemperature)
                    return
            except Exception:
                continue
        raise AttributeError(
            f"No writable temperature node found for block '{Blockname}'. Tried CTEMP, REAC_TEMP, TEMP."
        )

    def BLK_RADFRAC_Set_Refluxratio(self, Blockname, Refluxratio):
        self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_RR").Value = float(Refluxratio)

    def BLK_RADFRAC_Set_BoilupRatio(self, Blockname, BoilupRatio):
        self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_BR").Value = float(BoilupRatio)

    def _read_per_error(self):
        node = self.AspenSimulation.Tree.FindNode(r"\Data\Results Summary\Run-Status\Output\PER_ERROR")
        if node is None:
            raise RuntimeError("PER_ERROR node was not found.")
        value = getattr(node, "Value", None)
        if value is None:
            raise RuntimeError("PER_ERROR node returned None.")
        return int(value)

    def Run(self):
        self.last_run_error = None
        self.last_per_error = None
        self.DismissAspenDialogs()
        start_time = time.time()
        try:
            self.AspenSimulation.Engine.Run2()
            elapsed = time.time() - start_time
            print(f"Runtime = {elapsed}")
            self.DismissAspenDialogs()
            per_error = self._read_per_error()
            self.last_per_error = per_error
            print("per_error value : ", per_error)
        except com_error as err:
            self.last_run_error = f"Aspen COM error during run or PER_ERROR read: {err}"
            print(self.last_run_error)
            return False
        except Exception as err:
            self.last_run_error = f"Aspen run failed while executing or reading PER_ERROR: {err}"
            print(self.last_run_error)
            return False

        if per_error == 0:
            return True

        self.last_run_error = f"Aspen PER_ERROR was non-zero (PER_ERROR={self.last_per_error})."
        print("Aspen PER_ERROR was non-zero.")
        return False


def create_simulation(aspen_filename, working_directory, visible=False, suppress_dialogs=False):
    simulation = AspenSimulation(
        AspenFileName=aspen_filename,
        WorkingDirectoryPath=working_directory,
        VISIBILITY=visible,
        suppress_dialogs=suppress_dialogs,
    )
    print(f"Aspen visible: {getattr(simulation.AspenSimulation, 'Visible', 'unknown')}")
    print(f"Aspen dialog suppression: {suppress_dialogs}")
    print(f"Aspen document: {simulation.abs_path}")
    return simulation


def clamp01(x):
    return max(0.0, min(1.0, float(x)))


def normalize_utility_cost(cost, cost_ref=1000.0):
    cost = max(0.0, float(cost))
    cost_ref = max(1e-8, float(cost_ref))
    return np.log1p(cost) / np.log1p(cost_ref)


def compute_reward_v2(
    purity,
    selectivity,
    utility_cost,
    current_action,
    prev_action,
    w_purity=1.0,
    w_selectivity=1.0,
    w_balance=0.5,
    w_cost=0.7,
    w_action_penalty=0.0,
    utility_cost_ref=1000.0,
    purity_target=0.90,
    selectivity_target=0.30,
    target_steepness=10.0,
):
    purity = clamp01(purity)
    selectivity = clamp01(selectivity)

    quality_score = (w_purity * purity) + (w_selectivity * selectivity)
    balance_bonus = w_balance * np.sqrt(max(purity * selectivity, 0.0))

    purity_diff = purity - purity_target
    selectivity_diff = selectivity - selectivity_target
    target_score = 0.5 * (
        np.tanh(target_steepness * purity_diff) + np.tanh(target_steepness * selectivity_diff)
    )

    utility_penalty = w_cost * normalize_utility_cost(utility_cost, cost_ref=utility_cost_ref)

    action_penalty = 0.0
    if prev_action is not None:
        action_penalty = w_action_penalty * np.linalg.norm(current_action - prev_action)

    reward = quality_score + balance_bonus + target_score - utility_penalty - action_penalty
    reward_info = {
        "purity": float(purity),
        "selectivity": float(selectivity),
        "utility_cost": float(utility_cost),
        "quality_score": float(quality_score),
        "balance_bonus": float(balance_bonus),
        "target_score": float(target_score),
        "action_penalty": float(action_penalty),
        "utility_penalty": float(utility_penalty),
        "reward": float(reward),
    }
    return float(reward), reward_info


class SimpleCMAES:
    def __init__(self, mean, sigma=0.18, population_size=None, seed=None):
        self.mean = np.asarray(mean, dtype=np.float64)
        self.n_dim = self.mean.size
        self.sigma = float(sigma)
        self.rng = np.random.default_rng(seed)
        self.population_size = int(population_size or (4 + np.floor(3 * np.log(self.n_dim))))
        self.population_size = max(self.population_size, 4)
        self.mu = max(2, self.population_size // 2)
        weights = np.log(self.mu + 0.5) - np.log(np.arange(1, self.mu + 1))
        self.weights = weights / weights.sum()
        self.mu_eff = 1.0 / np.sum(self.weights ** 2)

        self.c_sigma = (self.mu_eff + 2.0) / (self.n_dim + self.mu_eff + 5.0)
        self.d_sigma = 1.0 + 2.0 * max(0.0, np.sqrt((self.mu_eff - 1.0) / (self.n_dim + 1.0)) - 1.0) + self.c_sigma
        self.c_c = (4.0 + self.mu_eff / self.n_dim) / (self.n_dim + 4.0 + 2.0 * self.mu_eff / self.n_dim)
        self.c1 = 2.0 / ((self.n_dim + 1.3) ** 2 + self.mu_eff)
        self.c_mu = min(
            1.0 - self.c1,
            2.0 * (self.mu_eff - 2.0 + 1.0 / self.mu_eff) / ((self.n_dim + 2.0) ** 2 + self.mu_eff),
        )
        self.chi_n = np.sqrt(self.n_dim) * (1.0 - 1.0 / (4.0 * self.n_dim) + 1.0 / (21.0 * self.n_dim ** 2))

        self.p_sigma = np.zeros(self.n_dim, dtype=np.float64)
        self.p_c = np.zeros(self.n_dim, dtype=np.float64)
        self.C = np.eye(self.n_dim, dtype=np.float64)
        self.B = np.eye(self.n_dim, dtype=np.float64)
        self.D = np.ones(self.n_dim, dtype=np.float64)
        self.invsqrtC = np.eye(self.n_dim, dtype=np.float64)
        self.generation = 0

    def _update_eigensystem(self):
        eigenvalues, eigenvectors = np.linalg.eigh(self.C)
        eigenvalues = np.maximum(eigenvalues, 1e-12)
        self.D = np.sqrt(eigenvalues)
        self.B = eigenvectors
        self.invsqrtC = self.B @ np.diag(1.0 / self.D) @ self.B.T

    def ask(self):
        self._update_eigensystem()
        zs = self.rng.standard_normal((self.population_size, self.n_dim))
        ys = zs @ (self.B * self.D).T
        xs = self.mean + self.sigma * ys
        return np.clip(xs, 0.0, 1.0)

    def tell(self, solutions, fitnesses):
        solutions = np.asarray(solutions, dtype=np.float64)
        fitnesses = np.asarray(fitnesses, dtype=np.float64)
        if solutions.ndim != 2 or solutions.shape[0] == 0:
            return

        order = np.argsort(fitnesses)
        solutions = solutions[order]
        available_mu = min(self.mu, solutions.shape[0])
        weights = self.weights[:available_mu]
        weights = weights / weights.sum()

        old_mean = self.mean.copy()
        top_solutions = solutions[:available_mu]
        self.mean = np.sum(top_solutions * weights[:, None], axis=0)
        y_w = (self.mean - old_mean) / max(self.sigma, 1e-12)

        self.p_sigma = (
            (1.0 - self.c_sigma) * self.p_sigma
            + np.sqrt(self.c_sigma * (2.0 - self.c_sigma) * self.mu_eff) * (self.invsqrtC @ y_w)
        )
        norm_p_sigma = np.linalg.norm(self.p_sigma)
        h_sigma = float(
            norm_p_sigma / np.sqrt(1.0 - (1.0 - self.c_sigma) ** (2.0 * (self.generation + 1))) / self.chi_n
            < (1.4 + 2.0 / (self.n_dim + 1.0))
        )

        self.p_c = (
            (1.0 - self.c_c) * self.p_c
            + h_sigma * np.sqrt(self.c_c * (2.0 - self.c_c) * self.mu_eff) * y_w
        )

        y_k = (top_solutions - old_mean) / max(self.sigma, 1e-12)
        rank_mu = np.zeros_like(self.C)
        for weight, y_vec in zip(weights, y_k):
            rank_mu += weight * np.outer(y_vec, y_vec)

        delta_h_sigma = (1.0 - h_sigma) * self.c_c * (2.0 - self.c_c)
        self.C = (
            (1.0 - self.c1 - self.c_mu + self.c1 * delta_h_sigma) * self.C
            + self.c1 * np.outer(self.p_c, self.p_c)
            + self.c_mu * rank_mu
        )
        self.C = 0.5 * (self.C + self.C.T)
        self.sigma *= np.exp((self.c_sigma / self.d_sigma) * (norm_p_sigma / self.chi_n - 1.0))
        self.sigma = float(np.clip(self.sigma, 0.01, 0.50))
        self.generation += 1


class KeroseneCMAESEvaluator:
    def __init__(
        self,
        simulation,
        simulation_factory=None,
        action_low=None,
        action_high=None,
        reward_report_scale=10.0,
        dialog_suppression=False,
    ):
        self.sim = simulation
        self.simulation_factory = simulation_factory
        self.dialog_suppression_enabled = bool(dialog_suppression)

        self.low_act = np.array(action_low if action_low is not None else [220.0, 1.50, 0.30], dtype=np.float32)
        self.high_act = np.array(action_high if action_high is not None else [260.0, 2.20, 0.85], dtype=np.float32)
        self.reward_report_scale = max(1e-8, float(reward_report_scale))

        self.target_stream = "KERO"
        self.product_streams = ["KERO", "DIESEL"]
        self.default_nominal_action = np.array([240.0, 1.80, 0.50], dtype=np.float32)

        self.w_purity = 1.15
        self.w_selectivity = 1.10
        self.w_balance = 0.60
        self.w_cost = 0.45
        self.w_action_penalty = 0.00
        self.utility_cost_ref = 30.0
        self.purity_target = 0.925
        self.selectivity_target = 0.860
        self.target_steepness = 35.0

        self.action_retry_sleep = 0.5
        self.rpc_restart_sleep = 2.0
        self.debug_enabled = True
        self.run_call_count = 0
        self.prepare_call_count = 0
        self.evaluation_call_count = 0
        self.state = np.zeros(3, dtype=np.float32)
        self.last_good_action = None

        if hasattr(self.sim, "DialogSuppression"):
            self.sim.DialogSuppression(self.dialog_suppression_enabled)
        self.nominal_action = self._load_nominal_action_from_simulation()

    def _debug_log(self, message):
        if not self.debug_enabled:
            return
        timestamp = time.strftime("%H:%M:%S")
        print(f"[AspenEval][{timestamp}] {message}")

    def _report_reward(self, reward):
        return float(reward) * self.reward_report_scale

    def _safe_float(self, value, fallback):
        try:
            return float(value)
        except (TypeError, ValueError):
            return float(fallback)

    def _is_rpc_error(self, err):
        text = str(err)
        return (
            "RPC server is unavailable" in text
            or "The RPC server is unavailable" in text
            or "-2147023174" in text
        )

    def _sanitize_action(self, action):
        action = np.asarray(action, dtype=np.float32)
        return np.clip(action, self.low_act, self.high_act).astype(np.float32)

    def _run_with_debug(self, context):
        self.run_call_count += 1
        run_id = self.run_call_count
        self._debug_log(f"Run #{run_id} start | context={context}")
        start_time = time.time()
        converged = self.sim.Run()
        elapsed = time.time() - start_time
        run_error = getattr(self.sim, "last_run_error", None)
        self._debug_log(
            f"Run #{run_id} end | context={context} | converged={converged} | elapsed={elapsed:.2f}s | last_run_error={run_error}"
        )
        return converged

    def _restart_simulation(self, context, err):
        if self.simulation_factory is None:
            message = f"RPC recovery requested during {context}, but no simulation factory is configured: {err}"
            print(message)
            return False, message

        print(f"Restarting Aspen session after {context}: {err}")
        try:
            if self.sim is not None:
                try:
                    self.sim.CloseAspen()
                except Exception as close_err:
                    print(f"Ignoring Aspen close error during restart: {close_err}")
        except Exception:
            pass

        time.sleep(self.rpc_restart_sleep)

        try:
            replacement = self.simulation_factory()
            if hasattr(replacement, "DialogSuppression"):
                replacement.DialogSuppression(self.dialog_suppression_enabled)
            self.sim = replacement
            self.nominal_action = self._load_nominal_action_from_simulation()
            self.state = np.zeros(3, dtype=np.float32)
            self.last_good_action = None
            print(f"Aspen session restarted successfully after {context}.")
            return True, None
        except Exception as restart_err:
            message = f"Aspen session restart failed after {context}: {restart_err}"
            print(message)
            return False, message

    def _load_nominal_action_from_simulation(self):
        nominal = self.default_nominal_action.copy()

        temperature_value = None
        temperature_node = None
        for input_name in ("CTEMP", "REAC_TEMP", "TEMP"):
            value = self.sim.read_block_input_value("R301", input_name)
            if value is not None:
                temperature_value = value
                temperature_node = input_name
                break
        if temperature_value is None:
            print("Could not read nominal reactor temperature from Aspen: no valid input node found.")
        else:
            nominal[0] = self._safe_float(temperature_value, nominal[0])
            self._debug_log(f"nominal read | block=R301 | node={temperature_node} | value={nominal[0]:.4f}")

        boilup_value = self.sim.read_block_input_value("D401", "BASIS_BR")
        if boilup_value is None:
            print("Could not read nominal D401 boilup ratio from Aspen: no valid input node found.")
        else:
            nominal[1] = self._safe_float(boilup_value, nominal[1])
            self._debug_log(f"nominal read | block=D401 | node=BASIS_BR | value={nominal[1]:.4f}")

        reflux_value = self.sim.read_block_input_value("D401", "BASIS_RR")
        if reflux_value is None:
            print("Could not read nominal D401 reflux ratio from Aspen: no valid input node found.")
        else:
            nominal[2] = self._safe_float(reflux_value, nominal[2])
            self._debug_log(f"nominal read | block=D401 | node=BASIS_RR | value={nominal[2]:.4f}")

        nominal = self._sanitize_action(nominal)
        print(f"Loaded nominal action from Aspen: {nominal.tolist()}")
        return nominal

    def _apply_action(self, action):
        safe_action = self._sanitize_action(action)
        self.sim.BLK_RPLUG_Set_T_SPEC_Constant_Temp("R301", float(safe_action[0]))
        self.sim.BLK_RADFRAC_Set_BoilupRatio("D401", float(safe_action[1]))
        self.sim.BLK_RADFRAC_Set_Refluxratio("D401", float(safe_action[2]))
        return safe_action

    def _get_process_metrics(self):
        try:
            self._debug_log("metric extraction start")
            self._debug_log("metric extraction -> purity start")
            purity = float(self.sim.get_kerosene_purity(self.target_stream, 8, 16))
            self._debug_log(f"metric extraction -> purity end | value={purity:.4f}")
            self._debug_log("metric extraction -> selectivity start")
            selectivity = float(self.sim.get_ft_selectivity(self.target_stream, self.product_streams, 8, 16))
            self._debug_log(f"metric extraction -> selectivity end | value={selectivity:.4f}")
            self._debug_log("metric extraction -> utility_cost start")
            utility_cost = float(self.sim.get_utility_cost())
            self._debug_log(f"metric extraction -> utility_cost end | value={utility_cost:.4f}")
            return purity, selectivity, utility_cost
        except Exception as err:
            self._debug_log(f"metric extraction failed | err={err}")
            raise RuntimeError(f"Process metric extraction failed: {err}") from err

    def prepare_evaluation(self, context, allow_unstable_nominal=True, allow_restart=True):
        self.prepare_call_count += 1
        prepare_id = self.prepare_call_count
        nominal_action = self.nominal_action.copy()
        self._debug_log(
            f"prepare #{prepare_id} start | context={context} | nominal={nominal_action.tolist()} | allow_unstable={allow_unstable_nominal}"
        )

        try:
            applied_action = self._apply_action(nominal_action)
            self._debug_log(f"prepare #{prepare_id} applied | action={applied_action.tolist()}")
            time.sleep(self.action_retry_sleep)
            converged = self._run_with_debug(f"{context}/nominal")
            if hasattr(self.sim, "DialogSuppression"):
                self.sim.DialogSuppression(self.dialog_suppression_enabled)

            run_error = getattr(self.sim, "last_run_error", None)
            info = {
                "converged": bool(converged),
                "com_error": False,
                "restored": bool(converged),
                "used_nominal_prepare": True,
                "nominal_action": applied_action.tolist(),
                "error_message": run_error or "",
            }

            if converged:
                self.state = np.zeros(3, dtype=np.float32)
                self.last_good_action = applied_action.copy()
                self._debug_log(f"prepare #{prepare_id} success | context={context}")
                return self.state.copy(), info

            self._debug_log(f"prepare #{prepare_id} unstable | context={context} | run_error={run_error}")
            if run_error and self._is_rpc_error(run_error) and allow_restart:
                restarted, restart_error = self._restart_simulation(context, run_error)
                if restarted:
                    return self.prepare_evaluation(context, allow_unstable_nominal=allow_unstable_nominal, allow_restart=False)
                info["com_error"] = True
                info["error_message"] = restart_error or run_error
                return self.state.copy(), info

            self.state = np.zeros(3, dtype=np.float32)
            if allow_unstable_nominal:
                print("Nominal prepare returned a non-converged state, but evaluator will continue to candidate run.")
                return self.state.copy(), info

            return self.state.copy(), info
        except Exception as err:
            self._debug_log(f"prepare #{prepare_id} exception | context={context} | err={err}")
            if self._is_rpc_error(err) and allow_restart:
                restarted, restart_error = self._restart_simulation(context, err)
                if restarted:
                    return self.prepare_evaluation(context, allow_unstable_nominal=allow_unstable_nominal, allow_restart=False)
                message = restart_error or str(err)
            else:
                message = str(err)
            self.state = np.zeros(3, dtype=np.float32)
            return self.state.copy(), {"converged": False, "com_error": True, "error_message": message}

    def _build_failed_result(
        self,
        action,
        elapsed,
        message,
        stage,
        com_error,
        state,
        failure_reward,
        applied_action=None,
    ):
        applied = self._sanitize_action(applied_action if applied_action is not None else action)
        current_state = np.asarray(state if state is not None else np.zeros(3, dtype=np.float32), dtype=np.float32)
        reward_raw = float(failure_reward)
        return {
            "state": current_state,
            "reward_raw": reward_raw,
            "reward_display": self._report_reward(reward_raw),
            "converged": False,
            "com_error": bool(com_error),
            "applied_action": applied,
            "elapsed": float(elapsed),
            "message": str(message),
            "stage": stage,
            "reward_info": {},
        }

    def evaluate_action(self, action, evaluation_index, is_validation=False, failure_reward=-50.0):
        self.evaluation_call_count += 1
        start_time = time.time()
        state_before, prepare_info = self.prepare_evaluation(
            context=f"eval-{evaluation_index}",
            allow_unstable_nominal=True,
            allow_restart=True,
        )

        if prepare_info.get("com_error"):
            elapsed = time.time() - start_time
            message = prepare_info.get("error_message", "Aspen prepare failed.")
            print(f"[CMA-ES] Eval {evaluation_index} prepare COM error | {message}")
            return self._build_failed_result(
                action=action,
                elapsed=elapsed,
                message=message,
                stage="prepare",
                com_error=True,
                state=state_before,
                failure_reward=failure_reward,
            )

        if not prepare_info.get("converged", True):
            message = prepare_info.get("error_message", "Nominal prepare was unstable.")
            print(f"[CMA-ES] Eval {evaluation_index} nominal prepare unstable | {message}")

        requested_action = self._sanitize_action(action)
        print(
            f"[CMA-ES] Eval {evaluation_index} start | validate={is_validation} "
            f"| action={requested_action[0]:.4f}/{requested_action[1]:.4f}/{requested_action[2]:.4f}"
        )

        try:
            applied_action = self._apply_action(requested_action)
        except Exception as err:
            elapsed = time.time() - start_time
            message = f"action setter failed: {err}"
            print(f"[CMA-ES] Eval {evaluation_index} setter error | {message}")
            return self._build_failed_result(
                action=requested_action,
                elapsed=elapsed,
                message=message,
                stage="candidate/setter",
                com_error=self._is_rpc_error(err),
                state=state_before,
                failure_reward=failure_reward,
            )

        time.sleep(self.action_retry_sleep)
        converged = self._run_with_debug(f"eval-{evaluation_index}/candidate")
        if hasattr(self.sim, "DialogSuppression"):
            self.sim.DialogSuppression(self.dialog_suppression_enabled)

        run_error = getattr(self.sim, "last_run_error", None)
        if not converged:
            elapsed = time.time() - start_time
            message = run_error or "Candidate action run did not converge."
            is_com_error = bool(run_error and self._is_rpc_error(run_error))
            if is_com_error:
                self._restart_simulation(f"candidate run eval-{evaluation_index}", message)
            print(f"[CMA-ES] Eval {evaluation_index} failed | stage=candidate/run | message={message}")
            return self._build_failed_result(
                action=requested_action,
                elapsed=elapsed,
                message=message,
                stage="candidate/run",
                com_error=is_com_error,
                applied_action=applied_action,
                state=state_before,
                failure_reward=failure_reward,
            )

        try:
            purity, selectivity, utility_cost = self._get_process_metrics()
        except Exception as err:
            elapsed = time.time() - start_time
            message = str(err)
            is_com_error = self._is_rpc_error(err)
            if is_com_error:
                self._restart_simulation(f"metric extraction eval-{evaluation_index}", err)
            print(f"[CMA-ES] Eval {evaluation_index} failed | stage=metrics | message={message}")
            return self._build_failed_result(
                action=requested_action,
                elapsed=elapsed,
                message=message,
                stage="metrics",
                com_error=is_com_error,
                applied_action=applied_action,
                state=state_before,
                failure_reward=failure_reward,
            )

        self.state = np.array(
            [
                purity,
                selectivity,
                normalize_utility_cost(utility_cost, cost_ref=self.utility_cost_ref),
            ],
            dtype=np.float32,
        )
        reward_raw, reward_info = compute_reward_v2(
            purity=purity,
            selectivity=selectivity,
            utility_cost=utility_cost,
            current_action=applied_action,
            prev_action=self.nominal_action,
            w_purity=self.w_purity,
            w_selectivity=self.w_selectivity,
            w_balance=self.w_balance,
            w_cost=self.w_cost,
            w_action_penalty=self.w_action_penalty,
            utility_cost_ref=self.utility_cost_ref,
            purity_target=self.purity_target,
            selectivity_target=self.selectivity_target,
            target_steepness=self.target_steepness,
        )
        reward_display = self._report_reward(reward_raw)
        elapsed = time.time() - start_time
        self.last_good_action = applied_action.copy()

        print(
            f"[CMA-ES] Eval {evaluation_index} done | reward={reward_display:.4f} "
            f"(raw={reward_raw:.4f}) | time={elapsed:.2f}s"
        )

        return {
            "state": self.state.copy(),
            "reward_raw": float(reward_raw),
            "reward_display": float(reward_display),
            "converged": True,
            "com_error": False,
            "applied_action": applied_action.copy(),
            "elapsed": float(elapsed),
            "message": "",
            "stage": "ok",
            "reward_info": reward_info,
        }


def validate_aspen_session(evaluator):
    print("[CMA-ES][preflight] validate_aspen_session start")
    state, info = evaluator.prepare_evaluation("preflight", allow_unstable_nominal=True, allow_restart=True)
    print(f"[CMA-ES][preflight] validate_aspen_session after prepare | info={info}")
    if info.get("com_error"):
        reason = info.get("error_message", "Aspen prepare returned a COM error before CMA-ES start.")
        raise RuntimeError(f"Aspen prepare failed: {reason}")
    if not info.get("converged", True):
        print("[CMA-ES][preflight] nominal prepare is unstable, but evaluator path is still alive.")
    return state, info


def run_cma_es_optimization(
    evaluator,
    max_evaluations,
    validation_evaluations,
    population_size,
    sigma_init,
    filename,
    seed=None,
    failure_reward=-50.0,
    failure_fitness=1.0e6,
    max_total_com_errors=12,
    max_consecutive_failures=3,
):
    csv_files = {
        "reward": f"report_rewards_{filename}.csv",
        "action": f"report_actions_{filename}.csv",
        "state": f"report_states_{filename}.csv",
        "time": f"report_runTime_{filename}.csv",
    }
    headers = {
        "reward": ["episode", "total_reward_raw", "total_reward_display"],
        "action": ["episode", "action"],
        "state": ["episode", "last_state"],
        "time": ["episode", "runTime_sec"],
    }
    for key, path in csv_files.items():
        with open(path, "w", newline="") as file:
            csv.writer(file).writerow(headers[key])

    def normalize_action(action):
        return np.clip(
            (np.asarray(action, dtype=np.float32) - evaluator.low_act) / (evaluator.high_act - evaluator.low_act),
            0.0,
            1.0,
        )

    def denormalize_action(action_unit):
        return evaluator.low_act + np.asarray(action_unit, dtype=np.float32) * (evaluator.high_act - evaluator.low_act)

    optimization_evaluations = max(1, int(max_evaluations) - int(validation_evaluations))
    validation_evaluations = max(0, int(validation_evaluations))
    optimizer = SimpleCMAES(
        mean=normalize_action(evaluator.nominal_action),
        sigma=sigma_init,
        population_size=population_size,
        seed=seed,
    )

    reward_history = []
    validation_records = []
    best_reward = -np.inf
    best_action = evaluator.nominal_action.copy()
    total_converged_steps = 0
    total_com_errors = 0
    consecutive_failures = 0
    action_labels = ["temperature", "boilup", "reflux"]
    action_span = np.asarray(evaluator.high_act - evaluator.low_act, dtype=np.float32)
    exact_bound_tolerance = np.array([0.02, 0.001, 0.001], dtype=np.float32)
    near_bound_tolerance = np.maximum(0.02 * action_span, np.array([0.50, 0.015, 0.010], dtype=np.float32))
    bound_hit_counts = {
        "low": np.zeros(evaluator.low_act.shape[0], dtype=np.int32),
        "high": np.zeros(evaluator.low_act.shape[0], dtype=np.int32),
        "near_low": np.zeros(evaluator.low_act.shape[0], dtype=np.int32),
        "near_high": np.zeros(evaluator.low_act.shape[0], dtype=np.int32),
    }

    print(
        f"Training config | max_evaluations={max_evaluations} | optimization={optimization_evaluations} "
        f"| validation={validation_evaluations} | population={optimizer.population_size} | sigma={optimizer.sigma:.3f} "
        f"| reward_report_scale={evaluator.reward_report_scale:.1f}x(log only)"
    )

    evaluation_index = 0
    while evaluation_index < optimization_evaluations:
        candidate_units = optimizer.ask()
        generation_solutions = []
        generation_fitnesses = []

        for candidate_unit in candidate_units:
            if evaluation_index >= optimization_evaluations:
                break

            action = denormalize_action(candidate_unit)
            episode = evaluation_index
            result = evaluator.evaluate_action(
                action=action,
                evaluation_index=episode,
                is_validation=False,
                failure_reward=failure_reward,
            )
            reward_history.append(result["reward_raw"])
            evaluation_index += 1

            if result["converged"]:
                consecutive_failures = 0
                total_converged_steps += 1
                if result["reward_raw"] > best_reward:
                    best_reward = float(result["reward_raw"])
                    best_action = result["applied_action"].copy()
                bound_hit_counts["low"] += (result["applied_action"] <= evaluator.low_act + exact_bound_tolerance).astype(np.int32)
                bound_hit_counts["high"] += (result["applied_action"] >= evaluator.high_act - exact_bound_tolerance).astype(np.int32)
                bound_hit_counts["near_low"] += (result["applied_action"] <= evaluator.low_act + near_bound_tolerance).astype(np.int32)
                bound_hit_counts["near_high"] += (result["applied_action"] >= evaluator.high_act - near_bound_tolerance).astype(np.int32)
                fitness = -result["reward_raw"]
            else:
                consecutive_failures += 1
                fitness = float(failure_fitness)
                if result["com_error"]:
                    total_com_errors += 1

            generation_solutions.append(normalize_action(result["applied_action"]))
            generation_fitnesses.append(fitness)

            action_str = re.sub(r"\s+", ",", str(result["applied_action"].tolist()).strip())
            with open(csv_files["reward"], "a", newline="") as file:
                csv.writer(file).writerow([episode, result["reward_raw"], result["reward_display"]])
            with open(csv_files["action"], "a", newline="") as file:
                csv.writer(file).writerow([episode, action_str])
            with open(csv_files["state"], "a", newline="") as file:
                csv.writer(file).writerow([episode, np.asarray(result["state"]).tolist()])
            with open(csv_files["time"], "a", newline="") as file:
                csv.writer(file).writerow([episode, result["elapsed"]])

            best_display = evaluator._report_reward(best_reward) if np.isfinite(best_reward) else float("nan")
            status_text = "OK" if result["converged"] else "FAIL"
            message_suffix = f" | Stage: {result['stage']} | Msg: {result['message']}" if not result["converged"] else ""
            print(
                f"Episode {episode} | Status: {status_text} | Total Reward: {result['reward_display']:.4f} "
                f"(raw={result['reward_raw']:.4f}) | Time: {result['elapsed']:.2f}s | Probe: False | Validate: False "
                f"| Act: {result['applied_action'][0]:.4f}/{result['applied_action'][1]:.4f}/{result['applied_action'][2]:.4f} "
                f"| BestEp: {best_display:.4f} | COM: {total_com_errors} | Sigma: {optimizer.sigma:.4f}{message_suffix}"
            )

            if total_com_errors >= int(max_total_com_errors) or consecutive_failures >= int(max_consecutive_failures):
                raise RuntimeError(
                    f"CMA-ES aborted after repeated Aspen failures "
                    f"(total_com_errors={total_com_errors}, consecutive_failures={consecutive_failures}). "
                    f"Last stage={result['stage']} | last_message={result['message']}"
                )

        if generation_solutions:
            optimizer.tell(np.asarray(generation_solutions), np.asarray(generation_fitnesses))

    if not np.isfinite(best_reward):
        print("Skipping validation because CMA-ES did not find any converged action.")
    else:
        for _ in range(validation_evaluations):
            episode = evaluation_index
            result = evaluator.evaluate_action(
                action=best_action,
                evaluation_index=episode,
                is_validation=True,
                failure_reward=failure_reward,
            )
            evaluation_index += 1
            validation_records.append(
                {
                    "episode": episode,
                    "reward_raw": result["reward_raw"],
                    "reward_display": result["reward_display"],
                    "action": result["applied_action"].copy(),
                    "converged": result["converged"],
                }
            )

            action_str = re.sub(r"\s+", ",", str(result["applied_action"].tolist()).strip())
            with open(csv_files["reward"], "a", newline="") as file:
                csv.writer(file).writerow([episode, result["reward_raw"], result["reward_display"]])
            with open(csv_files["action"], "a", newline="") as file:
                csv.writer(file).writerow([episode, action_str])
            with open(csv_files["state"], "a", newline="") as file:
                csv.writer(file).writerow([episode, np.asarray(result["state"]).tolist()])
            with open(csv_files["time"], "a", newline="") as file:
                csv.writer(file).writerow([episode, result["elapsed"]])

            message_suffix = f" | Stage: {result['stage']} | Msg: {result['message']}" if not result["converged"] else ""
            print(
                f"Episode {episode} | Status: {'OK' if result['converged'] else 'FAIL'} | Total Reward: {result['reward_display']:.4f} "
                f"(raw={result['reward_raw']:.4f}) | Time: {result['elapsed']:.2f}s | Probe: False | Validate: True "
                f"| Act: {result['applied_action'][0]:.4f}/{result['applied_action'][1]:.4f}/{result['applied_action'][2]:.4f} "
                f"| BestEp: {evaluator._report_reward(best_reward):.4f} | COM: {total_com_errors} | Sigma: {optimizer.sigma:.4f}{message_suffix}"
            )

    if np.isfinite(best_reward):
        print(
            f"Best action summary | reward={evaluator._report_reward(best_reward):.4f} "
            f"(raw={best_reward:.4f}) | action={best_action[0]:.4f}/{best_action[1]:.4f}/{best_action[2]:.4f}"
        )

    if validation_records:
        converged_validation_records = [record for record in validation_records if record["converged"]]
        validation_reward_array = np.asarray([record["reward_raw"] for record in validation_records], dtype=np.float32)
        validation_reward_display_array = np.asarray([record["reward_display"] for record in validation_records], dtype=np.float32)
        validation_action_array = np.asarray([record["action"] for record in validation_records], dtype=np.float32)
        validation_action_mean = validation_action_array.mean(axis=0)
        validation_action_std = validation_action_array.std(axis=0)
        validation_episode_text = ", ".join(f"{record['episode']}:{record['reward_display']:.4f}" for record in validation_records)
        print(
            f"Validation summary | count={len(validation_records)} | converged={len(converged_validation_records)} "
            f"| mean={validation_reward_display_array.mean():.4f} | std={validation_reward_display_array.std():.4f} "
            f"| min={validation_reward_display_array.min():.4f} | max={validation_reward_display_array.max():.4f} "
            f"| raw_mean={validation_reward_array.mean():.4f}"
        )
        print(
            f"Validation action mean/std | T={validation_action_mean[0]:.4f}+/-{validation_action_std[0]:.4f} "
            f"| B={validation_action_mean[1]:.4f}+/-{validation_action_std[1]:.4f} "
            f"| R={validation_action_mean[2]:.4f}+/-{validation_action_std[2]:.4f}"
        )
        print(f"Validation episodes | {validation_episode_text}")

    if total_converged_steps > 0:
        print(
            f"Boundary summary | total_converged={total_converged_steps} "
            f"| exact_low={bound_hit_counts['low'].tolist()} | exact_high={bound_hit_counts['high'].tolist()} "
            f"| near_low={bound_hit_counts['near_low'].tolist()} | near_high={bound_hit_counts['near_high'].tolist()}"
        )
        boundary_flags = []
        for idx, label in enumerate(action_labels):
            exact_low_rate = float(bound_hit_counts["low"][idx]) / float(total_converged_steps)
            exact_high_rate = float(bound_hit_counts["high"][idx]) / float(total_converged_steps)
            near_low_rate = float(bound_hit_counts["near_low"][idx]) / float(total_converged_steps)
            near_high_rate = float(bound_hit_counts["near_high"][idx]) / float(total_converged_steps)
            if exact_high_rate >= 0.10 or near_high_rate >= 0.25:
                boundary_flags.append(f"{label}:high exact={100.0 * exact_high_rate:.1f}% near={100.0 * near_high_rate:.1f}%")
            if exact_low_rate >= 0.10 or near_low_rate >= 0.25:
                boundary_flags.append(f"{label}:low exact={100.0 * exact_low_rate:.1f}% near={100.0 * near_low_rate:.1f}%")
        if boundary_flags:
            print("Boundary flags | " + " | ".join(boundary_flags))

    return reward_history
