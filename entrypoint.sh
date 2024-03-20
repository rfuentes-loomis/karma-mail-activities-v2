#!/bin/sh
set -e
echo 'adding hosts...'
echo "10.3.2.174 vmlinpgsqluat2.loomissayles.com" >> /etc/hosts
echo "10.3.2.23 vmlinarcproxyuat1.loomissayles.com" >> /etc/hosts
echo 'starting sshd...'
/usr/sbin/sshd
echo 'starting app...'
exec npm run start
#exec tail -f /dev/null
