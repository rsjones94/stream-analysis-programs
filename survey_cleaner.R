rm(list=ls())

# HOW TO USE THIS PROGRAM
# specify name of path of the .xlsx document to read (MUST be .xlsx)
# specify name of .xlsx document to output
# set the working directory (where the output will be written to)
# run as source

# RULES FOR USE
# the data should just be an xlsx file with one sheet, with headers (like you'd copy from the data collector's .txt file)
# once you begin a shot sequence (e.g., xs2) all following points must be part of that sequence until the sequence is finished
  # the exception to this is if you cut up a continuous profile with cross sections (or vice versa, but I don't know why that would happen)
# the character sequences thw, ws, bkf, hc, ri, po, and tob are protected (i.e., do not use these ANYWHERE in the shot name unless you are using them to designate a thalweg, etc.)
# the shots should be consistently named and of format NAME-designation1-designation2-...-designationN where - is a delimiter

# if you don't have a local copy of the package xlsx, install it by uncommenting the next line
#install.packages('xlsx')
library(xlsx)

# This script was last checked to work with R v3.3.0, 02/14/2018




########### USER INPUT
# note that filepaths for windows users use forward slashes, while Linux and Mac filepaths use backslashes

mywd = "X:\\2017\\161706201_WTSP\\Field\\Survey" # filepath to folder where your data is

datasource = 'WTSP_02-12-18' # filename of your input data (MUST be .xlsx)


########### END USER INPUT




































### create output names

outputname =  paste(datasource,'_clean',sep='')


delimit = '-' # the delimiter you used to separate the name of the shot (e.g., prot2) from its descriptors (e.g., thw)
myr = 'DEFUNCT' # what monitoring year is it? (strings are preferred)


prooutname = paste(outputname,'_pro.xlsx',sep='')
xsoutname = paste(outputname,'_xs.xlsx',sep='')


### read in the data and split it into profiles and x-sections
mywd = normalizePath(mywd)
setwd(mywd)

xcl = '_' # the character that excludes marker information from consideration for shot classification

datasource = paste(datasource,'.xlsx',sep='')
data = read.xlsx(datasource, sheetIndex = 1, header = TRUE, stringsAsFactors = FALSE)

allnames = data[,5]
namevec = NULL
for (i in 1:length(allnames)) {
  
  fullname = allnames[i] # grab the whole name e.g. prot1=thw=rock
  index = as.numeric(regexpr(delimit, fullname)) # get the index of the first delimiter
  endex = as.numeric(regexpr(xcl, fullname)) # get the index of the xcl char
  
  if (index != -1){
    name = substr(fullname,1,index-1) # get only the name up to the first delimiter e.g. prot1
  } else if (endex != -1) {
    name = substr(fullname,1,endex-1) # if no index, get name up to xcl char
  } else {
    name = fullname # in the case that there is no delimiter or xcl, make the name the whole string
  }
  
  rowname = name
  namevec[i] = rowname
  
}

pro = data[grep("pro", namevec, value = FALSE, ignore.case = TRUE),]
cross = data[grep("xs", namevec, value = FALSE, ignore.case = TRUE),]


### write a function marker() that gives the name of a profile or cross section given the row index of the list of either profiles or x-sections

marker = function(row, listname) {
  
  fullname = toString(listname[row,5]) # grab the whole name e.g. prot1=thw=rock # not suited for vectors of strings
  index = as.numeric(regexpr(delimit, listname[row,5])) # get the index of the first delimiter
  
  name = substr(fullname,1,index-1) # get only the name up to the first delimiter e.g. prot1
  
  # in the case that there is no delimiter, make the name the whole string
  if (name == '') {
    name = fullname
  }
  
  return(name)
  
}

### write a function projector() that finds the orthogonal projection of a onto b

projector = function(a,b) {
  
  maga = sqrt(sum(a^2))
  magb = sqrt(sum(b^2))
  
  unitb = b/magb
  costheta = a%*%b / (maga*magb)
  
  a1 = maga*costheta
  
  proj = a1*unitb
  
  return(proj)
  
}

### grab all XS shot sequences and write them to an XLSX file in individual worksheets
if(nrow(cross) != 0) {
# grab all the names for each row and make a vector
namevec = NULL
for (i in 1:nrow(cross)) {
  
  fullname = toString(cross[i,5]) # grab the whole name e.g. prot1=thw=rock
  index = as.numeric(regexpr(delimit, fullname)) # get the index of the first delimiter
  endex = as.numeric(regexpr(xcl, fullname)) # get the index of the xcl char
  
  if (index != -1){
    name = substr(fullname,1,index-1) # get only the name up to the first delimiter e.g. prot1
  } else if (endex != -1) {
    name = substr(fullname,1,endex-1) # if no index, get name up to xcl char
  } else {
    name = fullname # in the case that there is no delimiter or xcl, make the name the whole string
  }

  
  rowname = name
  namevec[i] = rowname
  
}

shotvec = NULL
for (i in 1:nrow(cross)) { # now get the shot description (after first delimiter, before the exclusionary char)
  
  fullname = toString(cross[i,5]) # grab the whole name e.g. prot1=thw=rock
  
  desdex = as.numeric(regexpr(delimit, fullname)) # get the index of the first delimiter
  endex = as.numeric(regexpr(xcl, fullname)) # get the index of the exclusionary character ( _ )
  
  
  if (endex == -1 & desdex == -1) { # no delimiter or exclusionary character
    
    endname = ''
    
  } else {
    
    if (endex == -1) { # if there is no delimiter
      endex = nchar(fullname)+1
    }
    if (desdex == -1) { # if there is no exclusionary character
      desdex = nchar(fullname)+1
    }
    
    endname = substr(fullname,desdex+1,endex-1) # get descriptive name
    
  }
  
  shotvec[i] = endname
  
}

# column bind this vector to the dataframe
ncross = cbind(cross,namevec)

# split the data by namevec
blankrow = rep(NA,ncol(ncross))

nncross = NULL
lastname = toString(ncross[1,ncol(ncross)])

for (i in 1:nrow(ncross)) {

  currentname = toString(ncross[i,ncol(ncross)])
  
  if (currentname != lastname | i == nrow(ncross)) { # if there is a change in name or we get to the end of the dataframe
    
    if (i == nrow(ncross)) { # if at the end of the dataframe
      nncross = rbind(nncross, ncross[i,])
    }
    
    # find the stationing
    origin = as.numeric(nncross[1,c(2,3)])
    modutmN = nncross[,2] - origin[1] # get the coordinates in meters relative to origin
    modutmE = nncross[,3] - origin[2]
    modutm = cbind(modutmN,modutmE)
    # note - there are three ways now to find the stationing
    # METHOD ONE - cumulative linear distance between each consecutive point (most error)
    # METHOD TWO - linear distance between each point and the origin
    # METHOD THREE - orthogonal projection onto the centerline between the origin and final point (least error)
    # this program uses the third method because why not
    # note that this is done algebraically
    lindist = sqrt(modutm[,1]^2 + modutm[,2]^2) # just used for sanity checking projected distance
    
    makevec = seq(from = 0, to = 0, length = nrow(modutm))
    projected = cbind(makevec,makevec)
    for (j in 2:nrow(modutm)) { # project all points onto the centerline
      
      projected[j,] = projector(modutm[j,],modutm[nrow(modutm),])
      
    }
    
    projdist = sqrt(projected[,1]^2 + projected[,2]^2)
    
    names(nncross)[ncol(nncross)] = 'Shot Sequence'
    
    nncross[,ncol(nncross)+1] = projdist
    names(nncross)[ncol(nncross)] = 'Station'
    
    thisyear = rep(NA,nrow(nncross))
    thisyear[1] = myr
    nncross[,ncol(nncross)+1] = thisyear
    names(nncross)[ncol(nncross)] = 'Monitoring Year'
    
    nr = nrow(nncross)
    
    bankfull = rep(NA,nr)
    nncross[,ncol(nncross)+1] = bankfull
    names(nncross)[ncol(nncross)] = 'Bankfull'
    
    xsname = rep(NA,nr)
    nncross[,ncol(nncross)+1] = xsname
    names(nncross)[ncol(nncross)] = 'XS Name'
    
    type = rep(NA,nr)
    nncross[,ncol(nncross)+1] = type
    names(nncross)[ncol(nncross)] = 'Reach Type (p/r)'
    
    reachdata = rep(NA,nr)
    nncross[,ncol(nncross)+1] = reachdata
    names(nncross)[ncol(nncross)] = 'Reach Membership'
    
    lowbank = rep(NA,nr)
    nncross[,ncol(nncross)+1] = reachdata
    names(nncross)[ncol(nncross)] = 'Low Bank Height'
    
    collectDate = rep(NA,nr)
    nncross[,ncol(nncross)+1] = reachdata
    names(nncross)[ncol(nncross)] = 'Collection Date'
    
    
    write.xlsx(x = nncross, file = xsoutname, sheetName = lastname, row.names = FALSE, showNA = FALSE, append = TRUE)
    nncross = ncross[i,]
    
  } else {
    
    nncross = rbind(nncross, ncross[i,]) # keep row
    
  }
  
  lastname = toString(ncross[i,ncol(ncross)])

}


}




### write each unique set of profile data to an XLSX file
if(nrow(pro) != 0) {
# grab all the names for each row and make a vector
namevec = NULL
for (i in 1:nrow(pro)) {
  
  fullname = toString(pro[i,5]) # grab the whole name e.g. prot1=thw=rock # not suited for vectors of strings
  index = as.numeric(regexpr(delimit, fullname)) # get the index of the first delimiter
  endex = as.numeric(regexpr(xcl, fullname)) # get the index of the xcl char
  
  if (index != -1){
    name = substr(fullname,1,index-1) # get only the name up to the first delimiter e.g. prot1
  } else if (endex != -1) {
    name = substr(fullname,1,endex-1) # if no index, get name up to xcl char
  } else {
    name = fullname # in the case that there is no delimiter or xcl, make the name the whole string
  }
  
  rowname = name
  namevec[i] = rowname
  
}

shotvec = NULL
for (i in 1:nrow(pro)) { # now get the shot description (after first delimiter, before the exclusionary char)
  
  fullname = toString(pro[i,5]) # grab the whole name e.g. prot1=thw=rock
  
  desdex = as.numeric(regexpr(delimit, fullname)) # get the index of the first delimiter
  endex = as.numeric(regexpr(xcl, fullname)) # get the index of the exclusionary character ( _ )
  
  
  if (endex == -1 & desdex == -1) { # no delimiter or exclusionary character
    
    endname = ''
    
  } else {
    
    if (endex == -1) { # if there is no delimiter
      endex = nchar(fullname)+1
    }
    if (desdex == -1) { # if there is no exclusionary character
      desdex = nchar(fullname)+1
    }
    
    endname = substr(fullname,desdex+1,endex-1) # get descriptive name
    
  }
  
  shotvec[i] = endname
  
}

# column bind this vector to the dataframe
npro = cbind(pro,namevec)

# split, clean and write data

lastname = toString(npro[1,ncol(npro)])
nnpro = NULL
en = 0 # initialize end
for (i in 1:nrow(npro)) {
  
  
  if (i == nrow(npro)) { # if you get to the end of the data, the way I wrote the next loop won't add the last row so you need this bit
    nnpro = rbind(nnpro,npro[i,])
  }
  
  currentname = toString(npro[i,ncol(npro)])
  
  if (currentname != lastname || i == nrow(npro)) { # if there is a change in name or we reach the end of the data
    
    # clean the data
    thw = rep(NA,nrow(nnpro))
    ws = rep(NA,nrow(nnpro))
    bkf = rep(NA,nrow(nnpro))
    tob = rep(NA,nrow(nnpro))
    xss = rep(NA,nrow(nnpro))
    str = rep(NA,nrow(nnpro))
    
    nnpro = cbind(nnpro,thw,ws,bkf,tob,xss,str) # add the six different shot types as columns
    
    bg = en + 1
    en = bg + nrow(nnpro) - 1 # end counter
    
    oldshotvec = NULL
    for (k in 1:nrow(nnpro)) {
      
      fullname = toString(nnpro[k,5]) # grab the whole name e.g. prot1=thw=rock
      
      desdex = as.numeric(regexpr(delimit, fullname)) # get the index of the first delimiter
      endex = as.numeric(regexpr(xcl, fullname)) # get the index of the exclusionary character ( _ )
      
      
      if (endex == -1 & desdex == -1) { # no delimiter or exclusionary character
        
        endname = ''
        
      } else {
        
        if (endex == -1) { # if there is no delimiter
          endex = nchar(fullname)+1
        }
        if (desdex == -1) { # if there is no exclusionary character
          desdex = nchar(fullname)+1
        }
        
        endname = substr(fullname,desdex+1,endex-1) # get descriptive name
        
      }
      
      oldshotvec[k] = endname
      
      
    }
    
    
    lastThw = 1
    for (j in 1:nrow(nnpro)) { # smoosh the data
      
      # we need to check what column this data should go in
      hasThw = length(grep("thw", oldshotvec[j], value = FALSE, ignore.case = TRUE))
      hasRi = length(grep("ri", oldshotvec[j], value = FALSE, ignore.case = TRUE))
      hasPo = length(grep("po", oldshotvec[j], value = FALSE, ignore.case = TRUE))
      hasWs = length(grep("ws", oldshotvec[j], value = FALSE, ignore.case = TRUE))
      hasBkf = length(grep("bkf", oldshotvec[j], value = FALSE, ignore.case = TRUE))
      hasTob = length(grep("tob", oldshotvec[j], value = FALSE, ignore.case = TRUE))
      hasHc = length(grep("hc", oldshotvec[j], value = FALSE, ignore.case = TRUE))
      hasXs = length(grep("xs", oldshotvec[j], value = FALSE, ignore.case = TRUE))
      hasStr = length(grep("st", oldshotvec[j], value = FALSE, ignore.case = TRUE))
      hasLog = length(grep("log", oldshotvec[j], value = FALSE, ignore.case = TRUE))
      hasLg = length(grep("lg", oldshotvec[j], value = FALSE, ignore.case = TRUE))
      

      
      if (hasThw == TRUE || hasRi == TRUE || hasPo == TRUE || hasHc == TRUE) {
        lastThw = j
        nnpro[lastThw,ncol(npro)+1] = nnpro[j,4]
      }
      
      if (hasWs == TRUE) {
        nnpro[lastThw,ncol(npro)+2] = nnpro[j,4]
      }
      
      if (hasBkf == TRUE) {
        nnpro[lastThw,ncol(npro)+3] = nnpro[j,4]
      }
      
      if (hasTob == TRUE) {
        nnpro[lastThw,ncol(npro)+4] = nnpro[j,4]
      }
      
      if (hasXs == TRUE) {
        nnpro[lastThw,ncol(npro)+5] = nnpro[j,4]
      }
      
      if (hasStr == TRUE | hasLog == TRUE | hasLg == TRUE) {
        nnpro[lastThw,ncol(npro)+6] = nnpro[j,4]
      }

      
    }
    
    for (j in nrow(nnpro):1){ # remove superfluous rows
      
      if (is.na(nnpro[j,ncol(npro)+1])) {
        
        nnpro = nnpro[-j,]
        
      }
      
    }
    
    XSCOPY = nnpro[,ncol(npro)+5]
    STRCOPY = nnpro[,ncol(npro)+6]
    
    nnpro = nnpro[,-(ncol(npro)+6)]  # remove the str column as the code did not originally have this column so it does not run when it is present
    nnpro = nnpro[,-(ncol(npro)+5)] # same for XSs
    nnpro = nnpro[,-4] # remove the elevation column as it is superfluous (and incomplete) as well
    
    
    # find the stationing
    origin = as.numeric(nnpro[1,c(2,3)])
    modutmN = nnpro[,2] - origin[1] # get the coordinates in meters relative to origin
    modutmE = nnpro[,3] - origin[2]
    modutm = cbind(modutmN,modutmE)
    
    # note - only way to find the stationing is cumulative distance
    delta = diff(modutm)
    dists = sqrt(delta[,1]^2+delta[,2]^2)
    cumdist = cumsum(dists)
    cumdist = c(0,cumdist)

    names(nnpro)[5] = 'Shot Sequence'
    
    nnpro[,ncol(nnpro)+1] = cumdist
    names(nnpro)[ncol(nnpro)] = 'Station'
    
    thisyear = rep(NA,nrow(nnpro))
    thisyear[1] = myr
    nnpro[,ncol(nnpro)+1] = thisyear
    names(nnpro)[ncol(nnpro)] = 'Monitoring Year'
    
    morph = rep(NA,nrow(nnpro))
    nnpro[,ncol(nnpro)+1] = morph
    names(nnpro)[ncol(nnpro)] = 'ri/ru/po/gl/st'
    
    proname = rep(NA,nrow(nnpro))
    nnpro[,ncol(nnpro)+1] = proname
    names(nnpro)[ncol(nnpro)] = 'Profile Name'
    
    reachdata = rep(NA,nrow(nnpro))
    nnpro[,ncol(nnpro)+1] = reachdata
    names(nnpro)[ncol(nnpro)] = 'Reach Membership'
    
    collectDate = rep(NA,nrow(nnpro))
    nnpro[,ncol(nnpro)+1] = reachdata
    names(nnpro)[ncol(nnpro)] = 'Collection Date'
    
    # guess ri/ru/po/gl/st
    
    newshotvec = NULL
    for (k in 1:nrow(nnpro)) {
      
      fullname = toString(nnpro[k,4]) # grab the whole name e.g. prot1=thw=rock
      
      desdex = as.numeric(regexpr(delimit, fullname)) # get the index of the first delimiter
      endex = as.numeric(regexpr(xcl, fullname)) # get the index of the exclusionary character ( _ )
      
      
      if (endex == -1 & desdex == -1) { # no delimiter or exclusionary character
        
        endname = ''
        
      } else {
        
        if (endex == -1) { # if there is no delimiter
          endex = nchar(fullname)+1
        }
        if (desdex == -1) { # if there is no exclusionary character
          desdex = nchar(fullname)+1
        }
        
        endname = substr(fullname,desdex+1,endex-1) # get descriptive name
        
      }
      
      newshotvec[k] = endname
      
      
    }
    
    isbpo = grep("bpo", newshotvec, value = FALSE, ignore.case = TRUE)
    isepo = grep("epo", newshotvec, value = FALSE, ignore.case = TRUE)
    isbri = grep("bri", newshotvec, value = FALSE, ignore.case = TRUE)
    iseri = grep("eri", newshotvec, value = FALSE, ignore.case = TRUE)
    

    
    nnpro[isbpo,12] = 'bpo'
    nnpro[isepo,12] = 'epo-bgl'
    nnpro[isbri,12] = 'bri'
    nnpro[iseri,12] = 'eri-bru'
    
    laster = ''
    for (m in 1:nrow(nnpro)) {
      
      isWhat = nnpro[m,12]
      if (is.na(isWhat)) {
        
      } else if (isWhat == 'bpo' || isWhat == 'bri') {
        nnpro[m,12] = paste(laster,'-',isWhat,sep='')
      } else if (isWhat == 'epo-bgl') {
        laster = 'egl'
      } else if (isWhat == 'eri-bru') {
        laster = 'eru'
      }
      
    }
    
    notNAs = which(!is.na(nnpro[,12]))
    firstelement = notNAs[1]
    lastelement = notNAs[length(notNAs)]
    
    if (!is.na(substr(nnpro[firstelement,12],2,4))){ # if there is ri/po data we can do ri-ru-po-gl assessments
      nnpro[firstelement,12] = substr(nnpro[firstelement,12],2,4)
      nnpro[lastelement,12] = substr(nnpro[lastelement,12],1,3)
    }

    
    blankprocol = rep(NA,nrow(nnpro))
    # reorder to dataframe into the new format
    nnpro = nnpro[,c(1,2,3,10,4,12,6,7,8,9)]
    nnpro = cbind(nnpro[,1:4],blankprocol,nnpro[,5:ncol(nnpro)],XSCOPY,STRCOPY)
    
    # make the simple morph vector
    compoundMorph = nnpro[,7]
    simpleMorph = blankprocol
    
    NonNAindex = which(!is.na(compoundMorph))
    firstNonNA = min(NonNAindex)
    firstCall = compoundMorph[firstNonNA]
    
    if (!is.na(firstCall)) {
      
      if(firstCall == 'bri') {
        prevCall = 'gl'
      } else if(firstCall == 'bpo') {
        prevCall = 'ru'
      }
      
      if(firstNonNA>1){
        simpleMorph[1:(firstNonNA-1)] = prevCall
      }
      
      for (k in firstNonNA:length(compoundMorph)) {
        
        isBRI = grepl('bri',compoundMorph[k])
        isERI = grepl('eri',compoundMorph[k])
        isRU = grepl('bru',compoundMorph[k])
        isBPO = grepl('bpo',compoundMorph[k])
        isEPO = grepl('epo',compoundMorph[k])
        isGL = grepl('bgl',compoundMorph[k])
        
        if(isBRI | isERI) {
          prevCall = 'ri'
        } else if(isBPO | isEPO) {
          prevCall = 'po'
        } else if(isRU) {
          prevCall = 'ru'
        } else if(isGL) {
          prevCall = 'gl'
        }
        
        simpleMorph[k] = prevCall
        
        if(isRU) {
          prevCall = 'ru'
        } else if(isGL) {
          prevCall = 'gl'
        }
        
      }
    }
    
    nnpro = cbind(nnpro[,1:7],simpleMorph,nnpro[,8:ncol(nnpro)])
    
    names(nnpro) = c('Shot Number', 'Northing', 'Easting', 'Station', 'Station+', 'Description', 'Compound.Morphology', 'Simplified.Morphology', 'thw', 'ws', 'bkf', 'tob', 'cross.section', 'structure')
    
    write.xlsx(x = nnpro, file = prooutname, sheetName = lastname, row.names = FALSE, append = TRUE, showNA = FALSE) # write the data
    
    if (i != nrow(npro)){
      nnpro = NULL # wipe nnpro
    }
    
  }
  
  if (i != nrow(npro)) {
    nnpro = rbind(nnpro,npro[i,])
  }
  
  lastname = toString(npro[i,ncol(npro)])
  
}


}
