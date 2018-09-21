/*rexx  */
/* Author: Andrew Jan                               */
/* Purpose: Transform an IOCDS deck to an Excel File*/
/* Latest update: 20170112                          */
/* CopyRight: IBM                                   */
/*arg pathi patho paths */

pathi = 'c:\IOCDS to XLS\Sample_IOCDS.txt'
patho = 'c:\IOCDS to XLS\Sample_IOCDS.xls'
pathe = 'c:\IOCDS to XLS\Sample_CPUCHIPs_SwitchPorts.txt'
/********************************************************/

filee = .stream~new(pathe)
ax. = ''
rec = filee~linein

do while length(rec) <> 0
  parse var rec ll ',' vv ';' rec
  ll = x2d(ll)
  ax.ll = vv
end
filee~close

patht1 = 'c:\IOCP\x.tmp'
patht2 = 'c:\IOCP\x.txt'
patht3 = 'c:\IOCP\x.xls'
filei = .stream~new(pathi)
fileo = .stream~new(patht2)
filet = .stream~new(patht1)
fileo~command("OPEN WRITE REPLACE")
filet~command("OPEN WRITE REPLACE")

/* merging the several lines of the same statement into one long line */
do while filei~lines <> 0
   rec = filei~linein
   if word(rec,1)='CHPID' | word(rec,1)='CNTLUNIT' | word(rec,1)='IODEVICE' then
     do
        xx = space(substr(rec,1,71))
        if substr(rec,72,1) \= ' ' then
          do until substr(rec,72,1) = ' '
            rec = filei~linein
            xx = xx || space(substr(rec,1,71))
          end
         /* say xx */
        filet~lineout(xx)
     end
end /* of do while thru the entire file */

filei~close
filet~close

/********************************************************/
ad. = ''
cu. = ''
adn. = ''
cun. = ''
unt. = ''
una. = ''
path. = ''
ln.  = ''
cuu. = ''
cua. = ''
prt. = ''
nprt. = ''
unta. = ''

i = 0
j = 0
k = 0

do while filet~lines <> 0
   rec = filet~linein
   if word(rec,1) = 'CHPID' then
     do
        j = j + 1
        MXCHP.j = rec
        MXCHPB.j = 0
     end
   else if word(rec,1) = 'CNTLUNIT' then
     do
        k = k + 1
        MXCNT.k = rec
        MXCNTB.k = 0
     end
   else if word(rec,1) = 'IODEVICE' then
     do
        parse var rec . 'ADDRESS=' addr ','
        if substr(addr,1,1) = '(' then
          do
            parse var rec . 'ADDRESS=' addr ')'
            addr = addr')'
          end
        parse var rec . 'CUNUMBR=(' cuno ')'
        parse var rec . 'PARTITION=((CSS(0),' part '))'
        parse var rec . 'NOTPART=((CSS(0),' nopart '))'
        parse var rec . 'UNIT=' unit ','
        parse var rec . 'UNITADD=' unta ','
        unit = space(unit)
        unta = space(unta)
        i = i + 1
        ad.i = addr
        cu.i = cuno
        if substr(addr,1,1) = '(' then
           do
             parse var addr '(' addr ','
           end
        adn.i = right(x2d(addr),5,'0')
        cun.i = right(x2d(cuno),5,'0')
        unt.i = unit
        una.i = unta
        prt.i = part
        nprt.i = nopart
     end
end /* of do while thru the entire file */

MXCHPB.0 = j
MXCNTB.0 = k

filet~close

tot = i

hit = 0

do i = 1 to tot

  do j = 1 to MXCNTB.0
    rec = MXCNT.j
    parse var rec . 'CUNUMBR=' cuno ',PATH=((CSS(0),' paths '))' .
    /* if cuno = cu.i then */
    if right(x2d(cuno),5,'0')= cun.i then
      do
        hit = 1
        MXCNTB.j = MXCNTB.j + 1
        path.i = paths
        parse var rec . 'UNITADD=' unita ','
        if substr(unita,1,1) = '(' then
          do
            parse var rec . 'UNITADD=(' unita ')'
            unita = unita')'
          end
        parse var rec . ',LINK=((CSS(0),' lnks '))'
        parse var rec . 'UNIT=' cunt ','
        parse var rec . 'CUADD=' cuad ','
        cunt = space(cunt)
        ln.i = lnks
        cuu.i = cunt
        cua.i = cuad
        unta.i = unita
        leave
      end
   end /* of do while thru the entire CNTLUNIT array */
  if hit = 0  then
    do
      say  '##' ad.i 'HAS NO CNTLUNIT ##'
    end
  hit = 0

end  /* of looping all devices againt all CNTLUNITs */


do i = 1 to tot

  do j = 1 to MXCHPB.0
     rec = MXCHP.j
     parse var rec . 'PATH=(CSS(0),' pp ')'
     if pos(pp,path.i) \= 0 then
       do
         hit = 1
         MXCHPB.j = MXCHPB.j + 1
         swt = ''
         parse var rec . 'PARTITION=' part
         y = lastpos(')',part) + 1
         parse var part part =(y)
         parse var rec . 'SWITCH=' swt ','
         parse var rec . 'PCHID=' pch ','
         parse var rec . 'TYPE=' typ
         typ = space(typ)
         xx = adn.i'|'ad.i'|'unt.i'|'una.i'|'prt.i'|'nprt.i'|'cun.i'|'cu.i'|'unta.i'|'
         if swt = '' then
            ent  = ''
         else
           do
            ll = x2d(pp)
            ent = ax.ll
           end
         xx = xx || path.i'|'ln.i'|'cua.i'|'cuu.i'|'right(x2d(pp),3,'0')'|'pp'|'right(x2d(pch),4,'0')'|'pch'|'swt'|'ent'|'typ'|'part'|'
         /* say xx */
         fileo~lineout(xx)
       end
   end /* of do while thru the entire CHIPID array */
   if  hit = 0  then
    do
      say  '##' ad.i cu.i 'HAS NO PATH ##'
    end
   hit = 0
end /* of looping all devices against paths */

fileo~close

/* the following lists those unused paths & Control Units */
  do j = 1 to MXCHPB.0
     xx = ''
     if MXCHPB.j = 0 then
      do
       xx = '##_NOT_IN_USE_## '
       say xx||MXCHPB.j MXCHP.j
      end
  end

  do j = 1 to MXCNTB.0
     xx = ''
     if MXCNTB.j = 0 then
      do
       xx = '##_NOT_IN_USE_## '
       say xx||MXCNTB.j MXCNT.j
      end
   end


address CMD 'cd c:\Program Files (x86)\Microsoft office\root\OFFICE16\'
address CMD 'EXCEL.exe /e "c:\IOCDS to XLS\IOCDS VBA2.xls"'

call SysFileDelete patht1   /* delete the temp file */

call SysFileDelete patht2   /* delete x.txt */

call SysFileCopy patht3, patho
call SysFileDelete patht3   /* delete x.txt */

exit 0   /* leave program */


