use crate::protocol::{
    CpuStats, DiskStats, GpuStats, HostInfo, MemoryStats, NetworkInterfaceStats, NetworkStats,
    PROTOCOL_VERSION, ProcessStats, TelemetrySnapshot,
};
use std::cmp::Ordering;
use std::process::Command;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use sysinfo::{Disks, Networks, System};

pub struct TelemetryCollector {
    system: System,
    disks: Disks,
    networks: Networks,
    sequence: u64,
    last_refresh: Instant,
    gpu: Option<GpuStats>,
    gpu_refresh_countdown: u8,
}

impl TelemetryCollector {
    pub fn new() -> Self {
        let mut system = System::new_all();
        system.refresh_all();
        Self {
            system,
            disks: Disks::new_with_refreshed_list(),
            networks: Networks::new_with_refreshed_list(),
            sequence: 0,
            last_refresh: Instant::now(),
            gpu: None,
            gpu_refresh_countdown: 0,
        }
    }

    pub fn snapshot(&mut self) -> TelemetrySnapshot {
        self.system.refresh_all();
        self.disks.refresh(true);
        self.networks.refresh(true);
        let elapsed = self.last_refresh.elapsed().as_secs_f64().max(0.001);
        self.last_refresh = Instant::now();
        self.sequence = self.sequence.saturating_add(1);

        if self.gpu_refresh_countdown == 0 {
            self.gpu = read_nvidia_gpu();
            self.gpu_refresh_countdown = 4;
        } else {
            self.gpu_refresh_countdown -= 1;
        }

        TelemetrySnapshot {
            protocol_version: PROTOCOL_VERSION,
            sequence: self.sequence,
            captured_at_unix_ms: SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap_or_default()
                .as_millis() as u64,
            host: self.host_info(),
            cpu: self.cpu_stats(),
            memory: self.memory_stats(),
            network: self.network_stats(elapsed),
            disks: self.disk_stats(),
            processes: self.process_stats(),
            gpu: self.gpu.clone(),
        }
    }

    fn host_info(&self) -> HostInfo {
        let cpu_model = self
            .system
            .cpus()
            .first()
            .map(|cpu| cpu.brand().trim().to_owned())
            .unwrap_or_else(|| "Unknown CPU".to_owned());
        HostInfo {
            user_name: std::env::var("USERNAME")
                .or_else(|_| std::env::var("USER"))
                .unwrap_or_else(|_| "Unknown user".to_owned()),
            hostname: System::host_name().unwrap_or_else(|| "Unknown host".to_owned()),
            os_name: System::name().unwrap_or_else(|| "Unknown OS".to_owned()),
            os_version: System::os_version().unwrap_or_default(),
            kernel_version: System::kernel_version().unwrap_or_default(),
            cpu_model,
            physical_cores: System::physical_core_count().unwrap_or(0),
            logical_cores: self.system.cpus().len(),
            uptime_seconds: System::uptime(),
        }
    }

    fn cpu_stats(&self) -> CpuStats {
        let cpus = self.system.cpus();
        let frequency_mhz = if cpus.is_empty() {
            0
        } else {
            cpus.iter().map(|cpu| cpu.frequency()).sum::<u64>() / cpus.len() as u64
        };
        CpuStats {
            usage_percent: self.system.global_cpu_usage(),
            frequency_mhz,
            per_core_percent: cpus.iter().map(|cpu| cpu.cpu_usage()).collect(),
        }
    }

    fn memory_stats(&self) -> MemoryStats {
        MemoryStats {
            total_bytes: self.system.total_memory(),
            used_bytes: self.system.used_memory(),
            available_bytes: self.system.available_memory(),
            swap_total_bytes: self.system.total_swap(),
            swap_used_bytes: self.system.used_swap(),
        }
    }

    fn network_stats(&self, elapsed: f64) -> NetworkStats {
        let mut interfaces = Vec::new();
        let mut received_bytes_per_second = 0_u64;
        let mut transmitted_bytes_per_second = 0_u64;
        let mut total_received_bytes = 0_u64;
        let mut total_transmitted_bytes = 0_u64;

        for (name, data) in &self.networks {
            let received = (data.received() as f64 / elapsed) as u64;
            let transmitted = (data.transmitted() as f64 / elapsed) as u64;
            received_bytes_per_second = received_bytes_per_second.saturating_add(received);
            transmitted_bytes_per_second = transmitted_bytes_per_second.saturating_add(transmitted);
            total_received_bytes = total_received_bytes.saturating_add(data.total_received());
            total_transmitted_bytes =
                total_transmitted_bytes.saturating_add(data.total_transmitted());
            interfaces.push(NetworkInterfaceStats {
                name: name.to_owned(),
                received_bytes_per_second: received,
                transmitted_bytes_per_second: transmitted,
                total_received_bytes: data.total_received(),
                total_transmitted_bytes: data.total_transmitted(),
                addresses: data
                    .ip_networks()
                    .iter()
                    .map(|address| format!("{}/{}", address.addr, address.prefix))
                    .collect(),
            });
        }

        interfaces.sort_by(|left, right| {
            let left_total = left.received_bytes_per_second + left.transmitted_bytes_per_second;
            let right_total = right.received_bytes_per_second + right.transmitted_bytes_per_second;
            right_total.cmp(&left_total)
        });

        NetworkStats {
            received_bytes_per_second,
            transmitted_bytes_per_second,
            total_received_bytes,
            total_transmitted_bytes,
            interfaces,
        }
    }

    fn disk_stats(&self) -> Vec<DiskStats> {
        let mut disks = self
            .disks
            .iter()
            .map(|disk| DiskStats {
                name: disk.name().to_string_lossy().into_owned(),
                mount_point: disk.mount_point().to_string_lossy().into_owned(),
                file_system: disk.file_system().to_string_lossy().into_owned(),
                kind: format!("{:?}", disk.kind()),
                total_bytes: disk.total_space(),
                available_bytes: disk.available_space(),
                removable: disk.is_removable(),
            })
            .collect::<Vec<_>>();
        disks.sort_by(|left, right| left.mount_point.cmp(&right.mount_point));
        disks
    }

    fn process_stats(&self) -> Vec<ProcessStats> {
        let logical_cores = self.system.cpus().len().max(1) as f32;
        let mut processes = self
            .system
            .processes()
            .iter()
            .map(|(pid, process)| {
                let disk = process.disk_usage();
                ProcessStats {
                    pid: pid.as_u32(),
                    name: process.name().to_string_lossy().into_owned(),
                    cpu_percent: process.cpu_usage() / logical_cores,
                    memory_bytes: process.memory(),
                    disk_read_bytes: disk.read_bytes,
                    disk_written_bytes: disk.written_bytes,
                    run_time_seconds: process.run_time(),
                }
            })
            .collect::<Vec<_>>();
        processes.sort_by(|left, right| {
            right
                .cpu_percent
                .partial_cmp(&left.cpu_percent)
                .unwrap_or(Ordering::Equal)
                .then_with(|| right.memory_bytes.cmp(&left.memory_bytes))
        });
        processes.truncate(80);
        processes
    }
}

impl Default for TelemetryCollector {
    fn default() -> Self {
        Self::new()
    }
}

fn read_nvidia_gpu() -> Option<GpuStats> {
    let mut command = Command::new("nvidia-smi");
    command.args([
        "--query-gpu=name,utilization.gpu,memory.used,memory.total,temperature.gpu",
        "--format=csv,noheader,nounits",
    ]);
    configure_hidden_command(&mut command);
    let output = command.output().ok()?;
    if !output.status.success() {
        return None;
    }
    let text = String::from_utf8(output.stdout).ok()?;
    let row = text.lines().next()?;
    let fields = row.split(',').map(str::trim).collect::<Vec<_>>();
    if fields.len() < 5 {
        return None;
    }
    Some(GpuStats {
        name: fields[0].to_owned(),
        usage_percent: fields[1].parse().ok()?,
        memory_used_bytes: fields[2].parse::<u64>().ok()?.saturating_mul(1024 * 1024),
        memory_total_bytes: fields[3].parse::<u64>().ok()?.saturating_mul(1024 * 1024),
        temperature_celsius: fields[4].parse().ok(),
    })
}

#[cfg(windows)]
fn configure_hidden_command(command: &mut Command) {
    use std::os::windows::process::CommandExt;
    command.creation_flags(0x08000000);
}

#[cfg(not(windows))]
fn configure_hidden_command(_: &mut Command) {}

pub async fn warm_up_collector(collector: &mut TelemetryCollector) {
    tokio::time::sleep(Duration::from_millis(250)).await;
    let _ = collector.snapshot();
}
