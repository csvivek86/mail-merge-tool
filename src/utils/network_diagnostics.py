"""
Network Diagnostics for NSNA Mail Merge Tool
Helper functions to diagnose network connectivity issues
"""

import socket
import smtplib
import ssl
import logging
from typing import Dict, List, Tuple

def test_network_connectivity() -> Dict[str, bool]:
    """Test basic network connectivity to common services"""
    results = {}
    
    # Test basic internet connectivity
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=10)
        results["internet"] = True
    except:
        results["internet"] = False
    
    # Test Gmail SMTP connectivity
    try:
        with socket.create_connection(("smtp.gmail.com", 587), timeout=15):
            results["gmail_smtp"] = True
    except:
        results["gmail_smtp"] = False
    
    # Test DNS resolution
    try:
        socket.gethostbyname("smtp.gmail.com")
        results["dns"] = True
    except:
        results["dns"] = False
    
    return results

def diagnose_smtp_connection(smtp_server: str, smtp_port: int) -> Tuple[bool, str]:
    """Diagnose SMTP connection issues and provide specific error information"""
    try:
        # Test basic socket connection
        with socket.create_connection((smtp_server, smtp_port), timeout=30) as sock:
            pass
        
        # Test SMTP handshake
        server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
        server.set_debuglevel(0)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.quit()
        
        return True, "SMTP connection successful"
        
    except socket.timeout:
        return False, f"Connection to {smtp_server}:{smtp_port} timed out. This may indicate:\n" \
                     "• Firewall blocking the connection\n" \
                     "• ISP blocking SMTP ports\n" \
                     "• Network connectivity issues"
    
    except socket.gaierror as e:
        return False, f"DNS resolution failed for {smtp_server}: {str(e)}\n" \
                     "• Check your internet connection\n" \
                     "• Verify DNS settings"
    
    except ConnectionRefusedError:
        return False, f"Connection refused by {smtp_server}:{smtp_port}\n" \
                     "• Server may be down\n" \
                     "• Port may be blocked"
    
    except ssl.SSLError as e:
        return False, f"SSL/TLS error: {str(e)}\n" \
                     "• Certificate issues\n" \
                     "• Outdated TLS version"
    
    except Exception as e:
        return False, f"Unexpected error: {str(e)}"

def get_network_diagnostic_report() -> str:
    """Generate a comprehensive network diagnostic report"""
    report = ["=== NSNA Mail Merge Tool - Network Diagnostics ===\n"]
    
    # Basic connectivity tests
    report.append("Basic Connectivity Tests:")
    connectivity = test_network_connectivity()
    
    report.append(f"  Internet: {'✓ OK' if connectivity.get('internet') else '✗ FAILED'}")
    report.append(f"  DNS Resolution: {'✓ OK' if connectivity.get('dns') else '✗ FAILED'}")
    report.append(f"  Gmail SMTP: {'✓ OK' if connectivity.get('gmail_smtp') else '✗ FAILED'}")
    report.append("")
    
    # SMTP-specific diagnostics
    report.append("SMTP Diagnostics:")
    smtp_ok, smtp_msg = diagnose_smtp_connection("smtp.gmail.com", 587)
    report.append(f"  Status: {'✓ OK' if smtp_ok else '✗ FAILED'}")
    report.append(f"  Details: {smtp_msg}")
    report.append("")
    
    # Recommendations
    if not connectivity.get('internet'):
        report.append("RECOMMENDATIONS:")
        report.append("• Check your internet connection")
        report.append("• Verify network settings")
        report.append("")
    
    if not connectivity.get('gmail_smtp') or not smtp_ok:
        report.append("SMTP TROUBLESHOOTING:")
        report.append("• Check firewall settings (allow port 587)")
        report.append("• Contact your ISP about SMTP blocking")
        report.append("• Try using a VPN")
        report.append("• Disable antivirus email scanning temporarily")
        report.append("• Try from a different network")
        report.append("")
    
    # System information
    import platform
    report.append("System Information:")
    report.append(f"  OS: {platform.system()} {platform.release()}")
    report.append(f"  Python: {platform.python_version()}")
    
    return "\n".join(report)

def show_connection_help() -> str:
    """Return helpful information for connection issues"""
    return """
Connection Timeout Solutions:

1. FIREWALL ISSUES:
   • Windows Defender: Add NSNA Mail Merge to allowed apps
   • Third-party antivirus: Disable email scanning temporarily
   • Corporate firewall: Contact IT to allow port 587

2. ISP BLOCKING:
   • Some ISPs block outgoing SMTP (port 25, 587)
   • Contact your ISP to confirm SMTP is allowed
   • Try using a VPN to bypass restrictions

3. NETWORK CONFIGURATION:
   • Proxy settings may interfere
   • Try from a different network (mobile hotspot)
   • Check router settings for port blocking

4. GOOGLE-SPECIFIC:
   • Ensure "Less secure app access" is configured if using app passwords
   • Verify OAuth credentials are valid
   • Check Google account security settings

5. ALTERNATIVE SOLUTIONS:
   • Use a different email provider
   • Configure through a different SMTP server
   • Export data and send manually

For persistent issues, run the network diagnostic tool from the Help menu.
"""
