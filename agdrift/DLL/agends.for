C**AGENDS
C  Continuum Dynamics, Inc.
C  Version 2.08 07/11/03
C
      SUBROUTINE AGENDS(IFLG,NNVEC,YYVEC,DDVEC)
!MS$ATTRIBUTES DLLEXPORT,STDCALL :: AGENDS
!MS$ATTRIBUTES REFERENCE :: IFLG
!MS$ATTRIBUTES REFERENCE :: NNVEC
!MS$ATTRIBUTES REFERENCE :: YYVEC
!MS$ATTRIBUTES REFERENCE :: DDVEC
C
C  AGENDS completes the solution for all desired output
C
C  IFLG   - Results flag: 0 to 28
C  NNVEC  - Number of points
C  YYVEC  - Distance array (m) [mostly]
C  DDVEC  -  0: Deposition (fraction applied)
C            1: Pond-integrated deposition (fraction applied)
C            2: Vertical flux (mg/sq cm)
C            3: One hour average concentration (ng/L)
C            4: COV
C            5: COV mean deposition (fraction applied)
C            6: Fraction aloft (fraction applied)
C            7: Single swath deposition (fraction applied)
C            8: Single swath upwind half boom deposition (fraction applied)
C            9: Multiple deposition (fraction applied)
C           10: Spray block deposition (fraction applied)
C           11: Canopy deposition (fraction)
C           12: Time accountancy aloft (fraction)
C           13: Time accountancy vapor (fraction)
C           14: Time accountancy canopy (fraction)
C           15: Time accountancy ground (fraction)
C           16: Height accountancy aloft (fraction)
C           17: Height accountancy vapor (fraction)
C           18: Height accountancy canopy (fraction)
C           19: Spray block drop size distribution
C           20: Downwind drop size distribution
C           21: Vertical flux drop size distribution
C           22: Distance accountancy aloft (fraction)
C           23: Distance accountancy vapor (fraction)
C           24: Distance accountancy canopy (fraction)
C           25: Distance accountancy ground (fraction)
C           26: Spray block area coverage (fraction)
C           27: Canopy drop size distribution
C           28: Application layout deposition (fraction applied)
C
      DIMENSION YYVEC(2),DDVEC(2)
C
      INCLUDE 'AGCOMMON.INC'
      INCLUDE 'AGSAMPLE.INC'
C
      COMMON /TEMP/ NTEMP,YTEMP(1620),ZTEMP(1620)
      COMMON /TBLK/ YN(4900),ZN(4900)
      COMMON /SBCV/ NSB,YSB(100),SSB(100),CSB(100)
      COMMON /SBCV/ NSBLK,FSBLK(201),CSBLK(201),FLXMX
      COMMON /APPL/ NBAP,NDAP,YBAP(201),ZBAP(201),YDAP(151),ZDAP(151)
C
C   0: Save the deposition results
C
      IF (IFLG.EQ.0) THEN
        DO N=1,NDEPS
          ZDEPS(N)=ZDEPT(N)
          ZDEPH(N)=ZDEPI(N)
        ENDDO
        CALL AGOVL
        NVEC=NDEPR
        NN=0
10      NN=NN+1
        IF (YDEPR(NN).LT.0.0) GOTO 10
        NNVEC=MIN0(NVEC-NN+1,NGRID)
        DO N=NN,NN+NNVEC-1
          YYVEC(N-NN+1)=YDEPR(N)
          DDVEC(N-NN+1)=ZDEPR(N)
          IF (YDEPR(N).GT.0.0.AND.YDEPR(N).LT.YGRID)
     $      ALEFT=ALEFT+0.5*DDEPR*(ZDEPR(N)+ZDEPR(N-1))
        ENDDO
        CALL AGAVE(NNVEC,YYVEC,DDVEC,NTEMP,YTEMP,ZTEMP)
        NNVEC=MIN0(NNVEC,NGRID-32)
C
        NDAP=MIN0(151,NNVEC)
        DO N=1,NDAP
          YDAP(N)=YYVEC(N)
          ZDAP(N)=DDVEC(N)
        ENDDO
C
C   1: Save the pond-integrated deposition
C
      ELSEIF (IFLG.EQ.1) THEN
        NNVEC=MIN0(NTEMP,NGRID-32)
        DO N=1,NNVEC
          YYVEC(N)=YTEMP(N)
          DDVEC(N)=ZTEMP(N)
        ENDDO
C
C   2: Save the flux results
C
      ELSEIF (IFLG.EQ.2) THEN
        FLXMX=0.0
        DO N=1,NFLXR
          ZFLXR(N)=ZFLXT(N)
          FLXMX=AMAX1(FLXMX,ZFLXT(N))
        ENDDO
        NVEC=NFLXR
        NN=NVEC
20      NN=NN-1
        IF (ZFLXR(NN).LT.0.001*FLXMX.AND.NN.GT.0) GOTO 20
        NNVEC=NN+1
        TEM=FLOW*AFRAC*DENN/UO/SWATH/0.1585
        DO N=1,NNVEC
          YYVEC(N)=YFLXR(N)
          DDVEC(N)=TEM*ZFLXR(N)
        ENDDO
C
C   3: Save the concentration results
C
      ELSEIF (IFLG.EQ.3) THEN
        NVEC=NFLXR
        NN=NVEC
30      NN=NN-1
        IF (ZFLXR(NN).LT.0.001*FLXMX.AND.NN.GT.0) GOTO 30
        NNVEC=NN
        TEM=1.0E+06*FLOW*AFRAC*DENN/UO/SWATH/0.1585/360.0
        DO N=1,NNVEC
          YYVEC(N)=YFLXR(N+1)
          IF (LCANF.EQ.0) THEN
            B=ALOG((YYVEC(N)+ZO)/ZO)
          ELSE
            IF (YYVEC(N).LE.HCAN) THEN
              B=UCAN*EXP(ALPHAC*(YYVEC(N)/HCAN-1.0))
            ELSE
              B=UOPN*ALOG((YYVEC(N)/HCAN-DOC+ZOC)/ZOC)
            ENDIF
          ENDIF
          UTEM=-SCW*B
          DDVEC(N)=TEM*ZFLXR(N+1)/UTEM
        ENDDO
C
C   4: Save the COV results
C
      ELSEIF (IFLG.EQ.4) THEN
        NSWTHH=MAX0(IFIX(80.0/SWATH)+1,NSWTH)
        NBLK=IFIX(NSWTHH*SWATH/2.0)+1
        SFAC=0.5*(1+IBOOM)
        IF (ISDTYP.NE.1) SFAC=SFAC+SDISP
        DO N=1,NBLK
          YN(N)=-2.0*(NBLK-N)
        ENDDO
        NSB=0
40      NSB=NSB+1
        SWATHN=SWATH*NSB/10
        NSWTHN=MAX0(IFIX(NSWTHH*SWATH/SWATHN),1)
        DO N=1,NBLK
          ZN(N)=0.0
        ENDDO
        NSWTM=NSWTHN
        IF (IBOOM.EQ.1) NSWTM=NSWTHN+1
        DO NS=1,NSWTM
          DO N=1,NBLK
            Y=YN(N)+(NS-SFAC)*SWATHN
            IF (NS.EQ.1.AND.IBOOM.EQ.1) THEN
              Z=0.5*AGINT(NDEPS,YDEPS,ZDEPH,Y)
            ELSEIF (NS.EQ.NSWTM.AND.IBOOM.EQ.1) THEN
              Z=0.5*(AGINT(NDEPS,YDEPS,ZDEPS,Y)
     $          -AGINT(NDEPS,YDEPS,ZDEPH,Y))
            ELSE
              Z=AGINT(NDEPS,YDEPS,ZDEPS,Y)
            ENDIF
            ZN(N)=ZN(N)+Z
          ENDDO
        ENDDO
        ZMEAN=0.0
        ZSPRD=0.0
        DO N=1,NBLK
          ZMEAN=ZMEAN+ZN(N)
          ZSPRD=ZSPRD+ZN(N)*ZN(N)
        ENDDO
        ZMEAN=ZMEAN/NBLK
        IF (ZMEAN.GT.0.0)
     $    ZSPRD=SQRT(AMAX1(0.0,ZSPRD/NBLK-ZMEAN*ZMEAN))/ZMEAN
        YSB(NSB)=SWATHN
        SSB(NSB)=ZSPRD
        CSB(NSB)=ZMEAN
        IF (NSB.LT.100.AND.SSB(NSB).LT.1.0) GOTO 40
        NNVEC=NSB
        DO N=1,NNVEC
          YYVEC(N)=YSB(N)
          DDVEC(N)=SSB(N)
        ENDDO
        SBCOV=SSB(10)
C
C   5: Save the COV mean deposition results
C
      ELSEIF (IFLG.EQ.5) THEN
        NNVEC=NSB
        DO N=1,NNVEC
          YYVEC(N)=YSB(N)
          DDVEC(N)=CSB(N)
        ENDDO
        SBMEAN=CSB(10)
C
C   6: Save the fraction aloft
C
      ELSEIF (IFLG.EQ.6) THEN
        NVEC=NDEPR
        NN=0
50      NN=NN+1
        IF (YDEPR(NN).LT.0.0) GOTO 50
        NNVEC=MIN0(NVEC-NN+1,NGRID)
        ATEM=ALEFT/SWATH/NSWTH+YDRFT/NSWTH
        DO N=NN,NN+NNVEC-1
          YYVEC(N-NN+1)=YDEPR(N)
          DDVEC(N-NN+1)=ATEM
          IF (YDEPR(N+1).LT.YGRID)
     $      ATEM=ATEM-0.5*DDEPR*(ZDEPR(N)+ZDEPR(N+1))/SWATH/NSWTH
        ENDDO
        NNVEC=MIN0(NNVEC,NGRID-32)
C
C   7: Save the single swath deposition results
C
      ELSEIF (IFLG.EQ.7) THEN
        NN=0
60      NN=NN+1
        IF (YDEPS(NN).LT.0.0) GOTO 60
        NNVEC=NDEPS-NN+1
        DO N=NN,NN+NNVEC-1
          YYVEC(N-NN+1)=YDEPS(N)
          DDVEC(N-NN+1)=ZDEPS(N)
        ENDDO
C
C   8: Save the single swath upwind half boom deposition
C
      ELSEIF (IFLG.EQ.8) THEN
        NN=0
70      NN=NN+1
        IF (YDEPS(NN).LT.0.0) GOTO 70
        NNVEC=NDEPS-NN+1
        DO N=NN,NN+NNVEC-1
          YYVEC(N-NN+1)=YDEPS(N)
          DDVEC(N-NN+1)=ZDEPH(N)
        ENDDO
C
C   9: Save the deposition results for multiple applications
C
      ELSEIF (IFLG.EQ.9) THEN
        DO N=1,NDEPS
          ZDEPS(N)=ZDEPT(N)
          ZDEPH(N)=ZDEPI(N)
        ENDDO
        CALL AGOVL
        NVEC=NDEPR
        NN=0
80      NN=NN+1
        IF (YDEPR(NN).LT.0.0) GOTO 80
        NDEPA=MIN0(NVEC-NN+1,NGRID)
        DO N=NN,NN+NDEPA-1
          YDEPA(N-NN+1)=YDEPR(N)
          ZDEPV(N-NN+1,LFMAC,LFMDR)=ZDEPR(N)
        ENDDO
        NDEPA=MIN0(NDEPA,NGRID-32)
        BLKSIZ=NSWTH*SWATH
C
C  10: Save the spray block deposition
C
      ELSEIF (IFLG.EQ.10) THEN
        DS=SWATH/(IFIX(SWATH/2.0)+1)
        SFAC=0.5*(1+IBOOM)
        NNVEC=IFIX((NSWTH-SDISP)*SWATH/DS)+1
        IF (ISDTYP.NE.1) SFAC=SFAC+SDISP
        SBLK=-(NSWTH-SDISP)*SWATH
        DO N=1,NNVEC
          YYVEC(N)=-DS*(NNVEC-N)
          DDVEC(N)=0.0
        ENDDO
        NSWTM=NSWTH
        IF (IBOOM.EQ.1) NSWTM=NSWTH+1
        DO NS=1,NSWTM
          DO N=1,NNVEC
            Y=YYVEC(N)+(NS-SFAC)*SWATH
            IF (NS.EQ.1.AND.IBOOM.EQ.1) THEN
              Z=0.5*AGINT(NDEPS,YDEPS,ZDEPH,Y)
            ELSEIF (NS.EQ.NSWTM.AND.IBOOM.EQ.1) THEN
              Z=0.5*(AGINT(NDEPS,YDEPS,ZDEPS,Y)
     $          -AGINT(NDEPS,YDEPS,ZDEPH,Y))
            ELSE
              Z=AGINT(NDEPS,YDEPS,ZDEPS,Y)
            ENDIF
            DDVEC(N)=DDVEC(N)+Z
          ENDDO
        ENDDO
C
        NBAP=MIN0(201,NNVEC)
        DO N=1,NBAP
          NN=N+NNVEC-NBAP
          YBAP(N)=YYVEC(NN)
          ZBAP(N)=DDVEC(NN)
        ENDDO
C
        TEM=0.0
        DO N=1,NNVEC
          TEM=AMAX1(TEM,DDVEC(N))
        ENDDO
        DSBLK=AMAX1(0.01,TEM/200)
        NSBLK=TEM/DSBLK+1
        DO NN=1,NSBLK
          FSBLK(NN)=DSBLK*(NN-1)
          CSBLK(NN)=0.0
        ENDDO
        DO N=2,NNVEC
          DTEMM=DDVEC(N-1)
          DTEMP=DDVEC(N)
          STEMM=YYVEC(N-1)
          STEMP=YYVEC(N)
          DO NN=1,10
            STEM=STEMM+(NN-0.5)*(STEMP-STEMM)/10
            IF (STEM.GT.SBLK) THEN
              DTEM=DTEMM+(NN-0.5)*(DTEMP-DTEMM)/10
              ITEM=DTEM/DSBLK+1
              CSBLK(ITEM)=CSBLK(ITEM)+1.0
            ENDIF
          ENDDO
        ENDDO
        CSBLK(1)=0.0
        DO NN=2,NSBLK
          CSBLK(NN)=CSBLK(NN)+CSBLK(NN-1)
        ENDDO
        IF (CSBLK(NSBLK).GT.0.0) THEN
          DO NN=1,NSBLK
            CSBLK(NN)=100.0*(1.0-CSBLK(NN)/CSBLK(NSBLK))
          ENDDO
        ENDIF
C
C  11: Save the canopy deposition
C
      ELSEIF (IFLG.EQ.11) THEN
        NNVEC=0
        TEMD=0.0
        DO N=NAHH,1,-1
          IF (TAHHV(N).LE.HCAN) THEN
            NNVEC=NNVEC+1
            YYVEC(N)=TAHHV(N)
            TEMD=TEMD+TAHFV(3,N)
            DDVEC(N)=1.0-TEMD
          ENDIF
        ENDDO
        CDEPS=TEMD
        IF (TEMD.EQ.0.0) NNVEC=0
C
C  12-15: Save the Time Accountancy results
C
      ELSEIF (IFLG.GE.12.AND.IFLG.LE.15) THEN
        NNVEC=NATT
        TEMD=0.0
        DO N=1,NNVEC
          YYVEC(N)=TATTV(N)
          IF (IFLG.EQ.12) THEN
            DDVEC(N)=1.0-TATFV(1,N)-TATFV(2,N)-TATFV(3,N)
          ELSEIF (IFLG.EQ.13) THEN
            DDVEC(N)=TATFV(1,N)
          ELSEIF (IFLG.EQ.14) THEN
            DDVEC(N)=TATFV(2,N)
          ELSE
            DDVEC(N)=TATFV(3,N)
          ENDIF
          TEMD=AMAX1(TEMD,DDVEC(N))
        ENDDO
        IF (TEMD.EQ.0.0) NNVEC=0
C
C  16-18: Save the Height Accountancy results
C
      ELSEIF (IFLG.GE.16.AND.IFLG.LE.18) THEN
        NNVEC=NAHH
        TEMD=0.0
        DO N=NNVEC,1,-1
          YYVEC(N)=TAHHV(N)
          IF (IFLG.EQ.16) THEN
            DDVEC(N)=1.0-TAHFV(1,N)-TAHFV(2,N)
            IF (N.GT.1) THEN
              TAHFV(1,N-1)=TAHFV(1,N-1)+TAHFV(1,N)
              TAHFV(2,N-1)=TAHFV(2,N-1)+TAHFV(2,N)
            ENDIF
          ELSEIF (IFLG.EQ.17) THEN
            DDVEC(N)=TAHFV(1,N)
          ELSE
            DDVEC(N)=TAHFV(2,N)
          ENDIF
          TEMD=AMAX1(TEMD,DDVEC(N))
        ENDDO
        IF (TEMD.EQ.0.0) NNVEC=0
C
C  19: Save the spray block drop size distribution
C
      ELSEIF (IFLG.EQ.19) THEN
        NNVEC=0.0
        TEM=0.0
        DO N=1,NDSD
          IF (DSSB(N).GT.0.00001) THEN
            TEM=TEM+DSSB(N)
            NNVEC=NNVEC+1
            YYVEC(NNVEC)=DSDC(N)
            DDVEC(NNVEC)=DSSB(N)
          ENDIF
        ENDDO
        IF (TEM.GT.0.0) THEN
          DO N=1,NNVEC
            DDVEC(N)=DDVEC(N)/TEM
          ENDDO
        ENDIF
C
C  20: Save the downwind drop size distribution
C
      ELSEIF (IFLG.EQ.20) THEN
        NNVEC=0.0
        TEM=0.0
        DO N=1,NDSD
          IF (DSDW(N).GT.0.00001) THEN
            TEM=TEM+DSDW(N)
            NNVEC=NNVEC+1
            YYVEC(NNVEC)=DSDC(N)
            DDVEC(NNVEC)=DSDW(N)
          ENDIF
        ENDDO
        IF (TEM.GT.0.0) THEN
          DO N=1,NNVEC
            DDVEC(N)=DDVEC(N)/TEM
          ENDDO
        ENDIF
C
C  21: Save the flux drop size distribution
C
      ELSEIF (IFLG.EQ.21) THEN
        NNVEC=0.0
        TEM=0.0
        DO N=1,NDSD
          IF (DSVP(N).GT.0.00001) THEN
            TEM=TEM+DSVP(N)
            NNVEC=NNVEC+1
            YYVEC(NNVEC)=DSDC(N)
            DDVEC(NNVEC)=DSVP(N)
          ENDIF
        ENDDO
        IF (TEM.GT.0.0) THEN
          DO N=1,NNVEC
            DDVEC(N)=DDVEC(N)/TEM
          ENDDO
        ENDIF
C
C  22-25: Save the Distance Accountancy results
C
      ELSEIF (IFLG.GE.22.AND.IFLG.LE.25) THEN
        NNVEC=NADD
        TEMD=0.0
        DO N=1,NNVEC
          YYVEC(N)=TADDV(N)
          IF (IFLG.EQ.22) THEN
            IF (N.GT.1) THEN
              DO I=1,3
                TADFV(I,N)=TADFV(I,N)+TADFV(I,N-1)
              ENDDO
            ENDIF
            DDVEC(N)=1.0-TADFV(1,N)-TADFV(2,N)-TADFV(3,N)
          ELSEIF (IFLG.EQ.23) THEN
            DDVEC(N)=TADFV(1,N)
          ELSEIF (IFLG.EQ.24) THEN
            DDVEC(N)=TADFV(2,N)
          ELSE
            DDVEC(N)=TADFV(3,N)
          ENDIF
          TEMD=AMAX1(TEMD,DDVEC(N))
        ENDDO
        IF (TEMD.EQ.0.0) NNVEC=0
C
C  26: Save the spray block area coverage
C
      ELSEIF (IFLG.EQ.26) THEN
        NNVEC=NSBLK
        DO N=1,NNVEC
          YYVEC(N)=FSBLK(N)
          DDVEC(N)=CSBLK(N)
        ENDDO
C
C  27: Save the canopy drop size distribution
C
      ELSEIF (IFLG.EQ.27) THEN
        NNVEC=0.0
        TEM=0.0
        DO N=1,NDSD
          IF (DSCP(N).GT.0.00001) THEN
            TEM=TEM+DSCP(N)
            NNVEC=NNVEC+1
            YYVEC(NNVEC)=DSDC(N)
            DDVEC(NNVEC)=DSCP(N)
          ENDIF
        ENDDO
        IF (TEM.GT.0.0) THEN
          DO N=1,NNVEC
            DDVEC(N)=DDVEC(N)/TEM
          ENDDO
        ENDIF
C
C  28: Save the application layout deposition
C
      ELSEIF (IFLG.EQ.28) THEN
        NNVEC=0
        DO N=1,NBAP
          IF (YBAP(N).GE.-200.0) THEN
            NNVEC=NNVEC+1
            YYVEC(NNVEC)=YBAP(N)
            DDVEC(NNVEC)=ZBAP(N)
          ENDIF
        ENDDO
        DO N=2,NDAP
          NNVEC=NNVEC+1
          YYVEC(NNVEC)=YDAP(N)
          DDVEC(NNVEC)=ZDAP(N)
        ENDDO
      ENDIF
      RETURN
      END