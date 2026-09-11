use serde::{Deserialize, Serialize};

pub const PROTOCOL_VERSION: u16 = 1;
pub const DEFAULT_PORT: u16 = 47_821;

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct TelemetrySnapshot {
    pub protocol_version: u16,
    pub sequence: u64,
    pub captured_at_unix_ms: u64,
    pub host: HostInfo,
    pub cpu: CpuStats,
    pub memory: MemoryStats,
    pub network: NetworkStats,
    pub disks: Vec<DiskStats>,
    pub processes: Vec<ProcessStats>,
    pub gpu: Option<GpuStats>,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct HostInfo {
    #[serde(default)]
    pub user_name: String,
    pub hostname: String,
    pub os_name: String,
    pub os_version: String,
    pub kernel_version: String,
    pub cpu_model: String,
    pub physical_cores: usize,
    pub logical_cores: usize,
    pub uptime_seconds: u64,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct CpuStats {
    pub usage_percent: f32,
    pub frequency_mhz: u64,
    pub per_core_percent: Vec<f32>,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct MemoryStats {
    pub total_bytes: u64,
    pub used_bytes: u64,
    pub available_bytes: u64,
    pub swap_total_bytes: u64,
    pub swap_used_bytes: u64,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct NetworkStats {
    pub received_bytes_per_second: u64,
    pub transmitted_bytes_per_second: u64,
    pub total_received_bytes: u64,
    pub total_transmitted_bytes: u64,
    pub interfaces: Vec<NetworkInterfaceStats>,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct NetworkInterfaceStats {
    pub name: String,
    pub received_bytes_per_second: u64,
    pub transmitted_bytes_per_second: u64,
    pub total_received_bytes: u64,
    pub total_transmitted_bytes: u64,
    pub addresses: Vec<String>,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct DiskStats {
    pub name: String,
    pub mount_point: String,
    pub file_system: String,
    pub kind: String,
    pub total_bytes: u64,
    pub available_bytes: u64,
    pub removable: bool,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct ProcessStats {
    pub pid: u32,
    pub name: String,
    pub cpu_percent: f32,
    pub memory_bytes: u64,
    pub disk_read_bytes: u64,
    pub disk_written_bytes: u64,
    pub run_time_seconds: u64,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct GpuStats {
    pub name: String,
    pub usage_percent: f32,
    pub memory_used_bytes: u64,
    pub memory_total_bytes: u64,
    pub temperature_celsius: Option<f32>,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
pub struct HealthResponse {
    pub status: String,
    pub service: String,
    pub protocol_version: u16,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum AgentCommand {
    TerminateProcess { request_id: u64, pid: u32 },
}

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum AgentEvent {
    ActionResult {
        request_id: u64,
        success: bool,
        message: String,
    },
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn snapshot_round_trip_preserves_protocol() {
        let snapshot = TelemetrySnapshot {
            protocol_version: PROTOCOL_VERSION,
            sequence: 42,
            host: HostInfo {
                hostname: "test-pc".to_owned(),
                ..HostInfo::default()
            },
            ..TelemetrySnapshot::default()
        };
        let json = serde_json::to_string(&snapshot).unwrap();
        let decoded: TelemetrySnapshot = serde_json::from_str(&json).unwrap();
        assert_eq!(decoded.protocol_version, PROTOCOL_VERSION);
        assert_eq!(decoded.sequence, 42);
        assert_eq!(decoded.host.hostname, "test-pc");
    }

    #[test]
    fn process_command_uses_tagged_wire_format() {
        let command = AgentCommand::TerminateProcess {
            request_id: 7,
            pid: 4242,
        };
        let json = serde_json::to_string(&command).unwrap();
        assert!(json.contains("terminate_process"));
        assert!(json.contains("4242"));
    }
}
