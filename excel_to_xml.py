from openpyxl import load_workbook


def to_xml():
    wb = load_workbook('摄像头ip.xlsx')
    ws = wb['Sheet1']
    i = 1
    str = ''
    for row in ws.rows:
        ip = row[0].value
        print(ip)
        res = echo_xml(ip)
        str += res
        i += 1
        if i > 2:
            break
    print(str)


def echo_xml(ip, name='摄像头', group='监控摄像头', securityname=1):
    str = """
    <host>
            <host>%s</host>
            <name>%s%s</name>
            <templates>
                <template>
                    <name>Template Module Generic SNMP</name>
                </template>
            </templates>
            <groups>
                <group>
                    <name>%s</name>
                </group>
            </groups>
            <interfaces>
                <interface>
                    <type>SNMP</type>
                    <ip>%s</ip>
                    <port>161</port>
                    <details>
                        <version>SNMPV3</version>
                        <securityname>%s</securityname>
                    </details>
                    <interface_ref>if1</interface_ref>
                </interface>
            </interfaces>
            <inventory_mode>DISABLED</inventory_mode>
        </host>
    """ % (ip, name, ip, group, ip, securityname)
    return str


if __name__ == '__main__':
    to_xml()
