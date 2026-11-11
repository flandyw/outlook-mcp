// Child-process test fixture only: simulate a Tailscale address on loopback.
import os from 'node:os';
import { syncBuiltinESMExports } from 'node:module';
const interfaces = os.networkInterfaces;
os.networkInterfaces = () => ({ ...interfaces(), tailscale0: [{
  address: '127.0.0.1', netmask: '255.0.0.0', family: 'IPv4',
  mac: '00:00:00:00:00:00', internal: false, cidr: '127.0.0.1/8',
}] });
syncBuiltinESMExports();
